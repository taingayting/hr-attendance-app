"""
HR worker book (stage 2): workers, documents checklist, photo and files.
Saved in the same database file as the attendance (memory/attendance.db). Nothing here touches your attendance Excel.
"""
import os, re, io, shutil
from datetime import datetime, date, timedelta

import attendance_db as _adb

WORKER_DIR = os.path.join("web_work", "workers")
ID_RE = re.compile(r"^DMO-\d{5}$")
DEFAULT_DOCS = ["CV", "ID card copy", "Photo 4x6", "Employment contract", "Health check paper", "Family book copy", "NSSF card"]
FIELDS = ["id", "dept", "title", "name_kh", "name_en", "entry_date", "sex", "married", "children", "union_member", "id_card",
          "province", "birthday", "nssf", "phone", "education", "stop_date", "stop_date_raw", "stop_kind", "stop_amount",
          "health_check", "remark", "note", "fp_registered"]
ALLOWED_FILES = (".pdf", ".jpg", ".jpeg", ".png", ".docx", ".xlsx", ".doc", ".xls")
MAX_FILE_MB = 10


def connect():
    con = _adb.connect()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS workers(
        id TEXT PRIMARY KEY, dept TEXT DEFAULT '', title TEXT DEFAULT '', name_kh TEXT DEFAULT '', name_en TEXT DEFAULT '',
        entry_date TEXT DEFAULT '', sex TEXT DEFAULT '', married TEXT DEFAULT '', children TEXT DEFAULT '', union_member TEXT DEFAULT '',
        id_card TEXT DEFAULT '', province TEXT DEFAULT '', birthday TEXT DEFAULT '', nssf TEXT DEFAULT '', phone TEXT DEFAULT '',
        education TEXT DEFAULT '', stop_date TEXT DEFAULT '', stop_date_raw TEXT DEFAULT '', stop_kind TEXT DEFAULT '', stop_amount TEXT DEFAULT '',
        health_check TEXT DEFAULT '', remark TEXT DEFAULT '', note TEXT DEFAULT '', created_at TEXT, updated_at TEXT,
        fp_registered TEXT DEFAULT '');
    CREATE TABLE IF NOT EXISTS doc_types(name TEXT PRIMARY KEY, sort INTEGER DEFAULT 0, required INTEGER DEFAULT 1);
    CREATE TABLE IF NOT EXISTS worker_docs(
        worker_id TEXT, doc_type TEXT, given INTEGER DEFAULT 0, given_on TEXT DEFAULT '', note TEXT DEFAULT '',
        PRIMARY KEY(worker_id, doc_type));
    CREATE TABLE IF NOT EXISTS contracts(
        id INTEGER PRIMARY KEY AUTOINCREMENT, worker_id TEXT, no INTEGER, start TEXT, end TEXT,
        made_on TEXT DEFAULT '', signed_on TEXT DEFAULT '', returned_on TEXT DEFAULT '', note TEXT DEFAULT '', created_at TEXT,
        UNIQUE(worker_id, no));
    CREATE TABLE IF NOT EXISTS holiday_paid(
        id INTEGER PRIMARY KEY AUTOINCREMENT, worker_id TEXT, name TEXT DEFAULT '', dept TEXT DEFAULT '', position TEXT DEFAULT '',
        day TEXT, holiday_name TEXT DEFAULT '', hours REAL DEFAULT 0, days REAL DEFAULT 0, amount REAL DEFAULT 0,
        date_paid TEXT DEFAULT '', method TEXT DEFAULT '', exclude TEXT DEFAULT 'YES', created_at TEXT,
        UNIQUE(worker_id, day));
    CREATE INDEX IF NOT EXISTS i_workers_dept ON workers(dept);
    CREATE TABLE IF NOT EXISTS recruit_requests(
        id INTEGER PRIMARY KEY AUTOINCREMENT, dept TEXT, position TEXT, need INTEGER DEFAULT 1, asked_by TEXT DEFAULT '',
        asked_on TEXT DEFAULT '', deadline TEXT DEFAULT '', note TEXT DEFAULT '', closed INTEGER DEFAULT 0, created_at TEXT);
    CREATE TABLE IF NOT EXISTS candidates(
        id INTEGER PRIMARY KEY AUTOINCREMENT, request_id INTEGER, name_en TEXT DEFAULT '', name_kh TEXT DEFAULT '', phone TEXT DEFAULT '',
        id_card TEXT DEFAULT '', sex TEXT DEFAULT '', found_on TEXT DEFAULT '', status TEXT DEFAULT 'Found', interview_on TEXT DEFAULT '',
        start_on TEXT DEFAULT '', note TEXT DEFAULT '', worker_id TEXT DEFAULT '', created_at TEXT);
    CREATE TABLE IF NOT EXISTS resignations(
        worker_id TEXT PRIMARY KEY, reason TEXT DEFAULT '', paper_done INTEGER DEFAULT 0, paper_on TEXT DEFAULT '',
        paid INTEGER DEFAULT 0, paid_on TEXT DEFAULT '', items_back INTEGER DEFAULT 0, items_on TEXT DEFAULT '',
        nssf_reported INTEGER DEFAULT 0, nssf_on TEXT DEFAULT '', note TEXT DEFAULT '');
    CREATE TABLE IF NOT EXISTS dept_plan(dept TEXT PRIMARY KEY, planned INTEGER DEFAULT 0);
    CREATE TABLE IF NOT EXISTS reminders(
        id INTEGER PRIMARY KEY AUTOINCREMENT, due TEXT, kind TEXT DEFAULT 'Reminder', worker_id TEXT DEFAULT '', text TEXT DEFAULT '',
        done INTEGER DEFAULT 0, created_at TEXT);
    CREATE TABLE IF NOT EXISTS warnings(
        id INTEGER PRIMARY KEY AUTOINCREMENT, worker_id TEXT, on_date TEXT, kind TEXT, reason TEXT DEFAULT '', created_at TEXT);
    CREATE TABLE IF NOT EXISTS snapshots(
        month TEXT PRIMARY KEY, working INTEGER, women INTEGER, men INTEGER, hired INTEGER, resigned INTEGER, depts TEXT, saved_at TEXT);
    CREATE TABLE IF NOT EXISTS factory_docs(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, kind TEXT DEFAULT '', expiry TEXT DEFAULT '', note TEXT DEFAULT '',
        file TEXT DEFAULT '', created_at TEXT);
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
    """)
    if "fp_registered" not in {r[1] for r in con.execute("PRAGMA table_info(workers)").fetchall()}:
        con.execute("ALTER TABLE workers ADD COLUMN fp_registered TEXT DEFAULT ''")      # '' or 'Y' = fingerprint works, 'N' = new worker, no fingerprint yet
    if "settled_on" not in {r[1] for r in con.execute("PRAGMA table_info(contracts)").fetchall()}:
        con.execute("ALTER TABLE contracts ADD COLUMN settled_on TEXT DEFAULT ''")       # the day the 5% end-of-contract money was paid (the column 结算)
    if not con.execute("SELECT 1 FROM doc_types LIMIT 1").fetchone():
        for i, n in enumerate(DEFAULT_DOCS):
            con.execute("INSERT OR IGNORE INTO doc_types(name, sort, required) VALUES(?,?,1)", (n, i))
    con.commit()
    return con


def _now():
    return datetime.now().isoformat(timespec="seconds")


def _rows(sql, args=()):
    con = connect()
    out = [dict(r) for r in con.execute(sql, args).fetchall()]
    con.close()
    return out


# ---------------------------------------------------------------- reading the CV sheet of your attendance Excel
def _txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return str(v).strip()


def parse_date(v):
    """Returns (iso 'YYYY-MM-DD' or '', raw text kept when it could not be read)."""
    if v is None or v == "":
        return "", ""
    if isinstance(v, datetime):
        return v.date().isoformat(), ""
    if isinstance(v, date):
        return v.isoformat(), ""
    s = str(v).strip()
    for f in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, f).date().isoformat(), ""
        except ValueError:
            pass
    return "", s


def read_cv_sheet(path):
    """Reads sheet CV. Returns (workers list of dict, warnings list). Read only: the Excel is never changed."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "CV" not in wb.sheetnames:
        wb.close()
        return [], ["This Excel has no sheet named CV."]
    ws = wb["CV"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    head = None
    for i, r in enumerate(rows[:6]):
        if r and len(r) > 1 and r[1] and "ID" in str(r[1]).upper() and "工号" in str(r[1]):
            head = i
            break
    if head is None:
        return [], ["I could not find the heading row (ID/工号) in the CV sheet."]
    out, warns, seen = [], [], set()
    for n, r in enumerate(rows[head + 1:], head + 2):
        r = list(r) + [None] * (26 - len(r))
        wid = _txt(r[1]).upper().replace(" ", "")
        if not wid:
            continue
        if not ID_RE.match(wid):
            warns.append(f"Row {n}: ID '{wid}' is not DMO- and 5 digits. Skipped.")
            continue
        if wid in seen:
            warns.append(f"Row {n}: {wid} is on the sheet twice. The second one is skipped.")
            continue
        seen.add(wid)
        entry, entry_raw = parse_date(r[6])
        birth, birth_raw = parse_date(r[13])
        stop, stop_raw = parse_date(r[18])
        if entry_raw: warns.append(f"{wid}: entry date '{entry_raw}' could not be read.")
        if birth_raw: warns.append(f"{wid}: birthday '{birth_raw}' could not be read.")
        out.append({
            "id": wid, "dept": _txt(r[2]).upper(), "title": _txt(r[3]), "name_kh": _txt(r[4]), "name_en": _txt(r[5]),
            "entry_date": entry, "sex": _txt(r[7]).upper(), "married": _txt(r[8]).upper(), "children": _txt(r[9]),
            "union_member": _txt(r[10]).upper(), "id_card": _txt(r[11]), "province": _txt(r[12]), "birthday": birth,
            "nssf": _txt(r[15]) if _txt(r[15]) != "0" else "", "phone": _txt(r[16]), "education": _txt(r[17]),
            "stop_date": stop, "stop_date_raw": stop_raw, "stop_kind": _txt(r[19]), "stop_amount": _txt(r[20]),
            "health_check": _txt(r[21]).upper(), "remark": " ".join(x for x in (_txt(r[22]), _txt(r[23])) if x), "note": ""})
    return out, warns


def import_preview(items):
    have = {w["id"] for w in _rows("SELECT id FROM workers")}
    new = [w for w in items if w["id"] not in have]
    return {"total": len(items), "new": len(new), "existing": len(items) - len(new)}


def import_commit(items):
    """Adds only the workers that are NOT in the worker book yet. Workers already in the book are never changed by an import."""
    con = connect()
    n = 0
    with con:
        have = {r[0] for r in con.execute("SELECT id FROM workers").fetchall()}
        for w in items:
            if w["id"] in have:
                continue
            cols = FIELDS + ["created_at", "updated_at"]
            con.execute(f"INSERT INTO workers({','.join(cols)}) VALUES({','.join('?' * len(cols))})",
                        [w.get(c, "") for c in FIELDS] + [_now(), _now()])
            n += 1
    con.close()
    return n


# ---------------------------------------------------------------- workers
def doc_types():
    return _rows("SELECT * FROM doc_types ORDER BY sort, name")


def add_doc_type(name):
    name = re.sub(r"\s+", " ", str(name or "")).strip()[:40]
    if not name:
        return False
    con = connect()
    with con:
        n = con.execute("SELECT COALESCE(MAX(sort), 0) + 1 FROM doc_types").fetchone()[0]
        con.execute("INSERT OR IGNORE INTO doc_types(name, sort, required) VALUES(?,?,1)", (name, n))
    con.close()
    return True


def list_workers():
    """All workers with the number of documents still missing (only the required document types count)."""
    ws = _rows("SELECT * FROM workers ORDER BY id")
    req = [d["name"] for d in doc_types() if d["required"]]
    given = {}
    for r in _rows("SELECT worker_id, doc_type FROM worker_docs WHERE given=1"):
        given.setdefault(r["worker_id"], set()).add(r["doc_type"])
    for w in ws:
        w["status"] = "Resigned" if (w["stop_date"] or w["stop_date_raw"]) else "Working"
        w["docs_missing"] = len([d for d in req if d not in given.get(w["id"], set())])
    return ws


def get_worker(wid):
    r = _rows("SELECT * FROM workers WHERE id=?", (wid,))
    return r[0] if r else None


def clean_worker(w):
    """Check what was typed. Returns (clean dict, error text)."""
    c = {f: re.sub(r"\s+", " ", str(w.get(f, "") or "")).strip() for f in FIELDS}
    c["id"] = c["id"].upper().replace(" ", "")
    if not ID_RE.match(c["id"]):
        return None, "The worker ID must look like DMO-10002 (DMO- and 5 digits)."
    c["dept"] = c["dept"].upper()
    if c["sex"] not in ("", "M", "F"):
        return None, "Sex must be M or F."
    for f in ("married", "union_member", "health_check", "fp_registered"):
        if c[f] not in ("", "Y", "N"):
            return None, f"{f.replace('_', ' ')} must be Y or N (or empty)."
    for f in ("entry_date", "birthday", "stop_date"):
        if c[f]:
            try:
                datetime.strptime(c[f], "%Y-%m-%d")
            except ValueError:
                return None, f"{f.replace('_', ' ')} is not a date."
    return c, ""


def save_worker(w, new=False):
    c, err = clean_worker(w)
    if err:
        return err
    con = connect()
    with con:
        exists = con.execute("SELECT 1 FROM workers WHERE id=?", (c["id"],)).fetchone()
        if new and exists:
            con.close()
            return f"{c['id']} is already in the worker book."
        if not new and not exists:
            con.close()
            return f"{c['id']} is not in the worker book."
        if new:
            cols = FIELDS + ["created_at", "updated_at"]
            con.execute(f"INSERT INTO workers({','.join(cols)}) VALUES({','.join('?' * len(cols))})",
                        [c[f] for f in FIELDS] + [_now(), _now()])
        else:
            sets = [f for f in FIELDS if f != "id"]
            con.execute(f"UPDATE workers SET {','.join(f + '=?' for f in sets)}, updated_at=? WHERE id=?",
                        [c[f] for f in sets] + [_now(), c["id"]])
    con.close()
    return ""


def age_of(birthday):
    try:
        b = datetime.strptime(birthday, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return ""
    t = date.today()
    return t.year - b.year - ((t.month, t.day) < (b.month, b.day))


# ---------------------------------------------------------------- documents checklist
def get_docs(wid):
    """One line for every document type, with what you ticked for this worker."""
    mine = {r["doc_type"]: r for r in _rows("SELECT * FROM worker_docs WHERE worker_id=?", (wid,))}
    out = []
    for d in doc_types():
        r = mine.get(d["name"], {})
        out.append({"doc_type": d["name"], "given": bool(r.get("given", 0)), "given_on": r.get("given_on", ""), "note": r.get("note", "")})
    return out


def save_docs(wid, items):
    con = connect()
    with con:
        for x in items:
            con.execute("INSERT OR REPLACE INTO worker_docs(worker_id, doc_type, given, given_on, note) VALUES(?,?,?,?,?)",
                        (wid, x["doc_type"], 1 if x["given"] else 0, x.get("given_on", "") if x["given"] else "", x.get("note", "")))
    con.close()


# ---------------------------------------------------------------- photo and files (saved in web_work/workers/<ID>/)
def worker_dir(wid):
    if not ID_RE.match(str(wid)):
        raise ValueError("bad worker id")
    return os.path.join(WORKER_DIR, wid)


def photo_path(wid):
    p = os.path.join(worker_dir(wid), "photo.jpg")
    return p if os.path.exists(p) else None


def save_photo(wid, data):
    """Check it is really a picture, make it at most 600 px, save as JPEG. Returns '' or an error text."""
    from PIL import Image, ImageOps
    try:
        im = Image.open(io.BytesIO(data))
        im.verify()
        im = Image.open(io.BytesIO(data))
        im = ImageOps.exif_transpose(im).convert("RGB")
    except Exception:
        return "That file is not a picture I can open (use JPG or PNG)."
    im.thumbnail((600, 600))
    os.makedirs(worker_dir(wid), exist_ok=True)
    im.save(os.path.join(worker_dir(wid), "photo.jpg"), "JPEG", quality=88)
    return ""


def safe_file_name(name):
    base = os.path.basename(str(name).replace("\\", "/"))
    base = re.sub(r"[^\w.\- ]", "_", base, flags=re.UNICODE).strip(" .")
    return base[:80] or "file"


def save_file(wid, doc_type, filename, data):
    """Save a scan or paper for this worker. Returns '' or an error text."""
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_FILES:
        return f"File type {ext or '(none)'} is not allowed. Use: " + ", ".join(ALLOWED_FILES)
    if len(data) > MAX_FILE_MB * 1024 * 1024:
        return f"The file is bigger than {MAX_FILE_MB} MB."
    d = os.path.join(worker_dir(wid), "files")
    os.makedirs(d, exist_ok=True)
    stem = safe_file_name(os.path.splitext(filename)[0])
    tag = safe_file_name(doc_type)
    path = os.path.join(d, f"{tag}__{datetime.now():%Y%m%d_%H%M%S}__{stem}{ext}")
    with open(path, "wb") as f:
        f.write(data)
    return ""


def list_files(wid):
    d = os.path.join(worker_dir(wid), "files")
    if not os.path.isdir(d):
        return []
    out = []
    for n in sorted(os.listdir(d), reverse=True):
        p = os.path.join(d, n)
        if os.path.isfile(p):
            out.append({"name": n, "path": p, "doc": n.split("__")[0], "size_kb": round(os.path.getsize(p) / 1024)})
    return out


# ---------------------------------------------------------------- contracts: a new 3-month contract again and again
CONTRACT_MONTHS = 3


def add_months(d, n):
    import calendar
    y, m = divmod(d.year * 12 + d.month - 1 + n, 12)
    m += 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


PROBATION_MONTHS = 2


def _first_of(d, n):
    """First day of the month that is n months after the month of d."""
    y, m = divmod(d.year * 12 + d.month - 1 + n, 12)
    return date(y, m + 1, 1)


def probation_end(entry):
    """Probation = the month the worker entered + the next month. Returns its last day."""
    return _first_of(entry, PROBATION_MONTHS) - timedelta(days=1)


def period(entry, no):
    """Contract number `no` (1, 2, 3 ...): whole calendar months. Contract 1 starts on the 1st of the 3rd month after the entry month
    (entered 27 Jun: probation Jun-Jul, contract 1 = 1 Aug to 31 Oct). Each contract is 3 months."""
    s = _first_of(entry, PROBATION_MONTHS + CONTRACT_MONTHS * (no - 1))
    e = _first_of(entry, PROBATION_MONTHS + CONTRACT_MONTHS * no) - timedelta(days=1)
    return s, e


def period_no_on(entry, on):
    """Which contract number is running on date `on` (0 while the worker is still in probation)."""
    if on < period(entry, 1)[0]:
        return 0
    no = 1
    while period(entry, no)[1] < on:
        no += 1
    return no


def contracts_of(wid=None):
    if wid:
        return _rows("SELECT * FROM contracts WHERE worker_id=? ORDER BY no DESC", (wid,))
    return _rows("SELECT * FROM contracts ORDER BY worker_id, no")


def save_contract(c):
    """c = dict(worker_id, no, start, end, made_on, signed_on, returned_on, note). Returns '' or an error text."""
    wid = str(c.get("worker_id", ""))
    if not ID_RE.match(wid):
        return "bad worker id"
    try:
        s, e = datetime.strptime(c["start"], "%Y-%m-%d").date(), datetime.strptime(c["end"], "%Y-%m-%d").date()
    except Exception:
        return f"{wid}: start and end must be dates."
    if e < s:
        return f"{wid}: the end date is before the start date."
    for f in ("made_on", "signed_on", "returned_on", "settled_on"):
        if c.get(f):
            try:
                datetime.strptime(c[f], "%Y-%m-%d")
            except ValueError:
                return f"{wid}: {f} is not a date."
    if c.get("signed_on") and not c.get("made_on"):
        return f"{wid}: it is signed but not made. Type the made date too."
    if c.get("returned_on") and not c.get("signed_on"):
        return f"{wid}: it is given back but not signed. Type the signed date too."
    con = connect()
    with con:
        if "settled_on" not in c:                             # a save that does not know the settlement date must not erase it
            old = con.execute("SELECT settled_on FROM contracts WHERE worker_id=? AND no=?", (wid, int(c["no"]))).fetchone()
            c = {**c, "settled_on": (old[0] if old else "") or ""}
        con.execute("INSERT INTO contracts(worker_id, no, start, end, made_on, signed_on, returned_on, note, created_at, settled_on) VALUES(?,?,?,?,?,?,?,?,?,?) "
                    "ON CONFLICT(worker_id, no) DO UPDATE SET start=excluded.start, end=excluded.end, made_on=excluded.made_on, "
                    "signed_on=excluded.signed_on, returned_on=excluded.returned_on, note=excluded.note, settled_on=excluded.settled_on",
                    (wid, int(c["no"]), c["start"], c["end"], c.get("made_on", ""), c.get("signed_on", ""), c.get("returned_on", ""),
                     str(c.get("note", ""))[:200], _now(), c.get("settled_on", "")))
    con.close()
    return ""


def delete_contract(cid):
    con = connect()
    with con:
        con.execute("DELETE FROM contracts WHERE id=?", (int(cid),))
    con.close()


def contract_rows(today=None, within_days=14):
    """One line for each working worker whose contract needs attention. Suggested periods are counted from the entry date;
    nothing is saved until you type a date. Returns list of dicts."""
    today = today or date.today()
    ws = [w for w in list_workers() if w["status"] == "Working" and w["entry_date"]]
    stored = {}
    for c in contracts_of():
        stored.setdefault(c["worker_id"], {})[c["no"]] = c
    out = []
    for w in ws:
        entry = datetime.strptime(w["entry_date"], "%Y-%m-%d").date()
        cur = period_no_on(entry, today)
        mine = stored.get(w["id"], {})
        cands = []
        if cur >= 1:
            cands.append(cur)
            cands.append(cur + 1)
        else:
            cands.append(1)
        # always show a stored contract that is not finished (made but not signed, signed but not given back)
        for no, c in mine.items():
            if c["made_on"] and not c["returned_on"] and no not in cands:
                cands.append(no)
        for no in sorted(set(cands)):
            c = mine.get(no)
            if c:
                s, e = datetime.strptime(c["start"], "%Y-%m-%d").date(), datetime.strptime(c["end"], "%Y-%m-%d").date()
            else:
                s, e = period(entry, no)
            done = bool(c and c["returned_on"])
            if done:
                continue
            if c and not c["made_on"] and not c["signed_on"] and not c["returned_on"] and no != cur + 1 and no != cur:
                continue
            days_to_start = (s - today).days
            days_to_end = (e - today).days
            if c:                                   # a stored contract that is still open
                state = "open"
            elif no == cur + 1 and days_to_start <= within_days:
                state = "due"                       # the next contract starts soon and is not made yet
            elif no == cur and mine and (cur - 1) in mine:
                state = "late"                      # you track this worker and the running contract was never made
            elif no == 1 and cur == 0 and days_to_start <= within_days:
                state = "due"
            else:
                continue
            out.append({"worker_id": w["id"], "name": w["name_en"], "name_kh": w["name_kh"], "dept": w["dept"], "title": w["title"],
                        "no": no, "start": s, "end": e, "made_on": (c or {}).get("made_on", ""), "signed_on": (c or {}).get("signed_on", ""),
                        "returned_on": (c or {}).get("returned_on", ""), "settled_on": (c or {}).get("settled_on", ""), "note": (c or {}).get("note", ""), "id": (c or {}).get("id"),
                        "entry": entry, "probation_end": probation_end(entry),
                        "state": state, "days_to_start": days_to_start, "days_to_end": days_to_end, "stored": bool(c)})
    out.sort(key=lambda r: (r["start"], r["worker_id"]))
    return out


def settlement_rows(today=None, within_days=30):
    """Contracts that END soon (or ended) and whose 5% end-of-contract money (the column 结算) is not marked as paid.
    Shows the running contract and the one before it (only when it ended in the last 60 days, or you stored it)."""
    today = today or date.today()
    stored = {}
    for c in contracts_of():
        stored.setdefault(c["worker_id"], {})[c["no"]] = c
    out = []
    for w in list_workers():
        if w["status"] != "Working" or not w["entry_date"]:
            continue
        entry = datetime.strptime(w["entry_date"], "%Y-%m-%d").date()
        cur = period_no_on(entry, today)
        mine = stored.get(w["id"], {})
        for no in sorted({n for n in (cur - 1, cur) if n >= 1} | {n for n in mine if mine[n]["made_on"]}):
            c = mine.get(no)
            if c and c["settled_on"]:
                continue
            s_, e_ = (datetime.strptime(c["start"], "%Y-%m-%d").date(), datetime.strptime(c["end"], "%Y-%m-%d").date()) if c else period(entry, no)
            days_left = (e_ - today).days
            if days_left > within_days or (not c and days_left < -7):       # an old contract you never stored: I cannot know if it was paid, so I do not nag
                continue
            out.append({"worker_id": w["id"], "name": w["name_en"], "dept": w["dept"], "no": no, "start": s_, "end": e_, "days_left": days_left,
                        "id": (c or {}).get("id"), "stored": bool(c)})
    out.sort(key=lambda r: (r["end"], r["worker_id"]))
    return out


def month_todo(year, month):
    """What you must do in one month: contracts that START in the month (make + sign), contracts that END in the month (pay the 5%),
    and workers whose probation ends. Returns three lists of dicts."""
    first = date(year, month, 1)
    last = _first_of(first, 1) - timedelta(days=1)
    stored = {}
    for c in contracts_of():
        stored.setdefault(c["worker_id"], {})[c["no"]] = c
    make, settle, prob = [], [], []
    for w in list_workers():
        if w["status"] != "Working" or not w["entry_date"]:
            continue
        entry = datetime.strptime(w["entry_date"], "%Y-%m-%d").date()
        mine = stored.get(w["id"], {})
        base = {"worker_id": w["id"], "name": w["name_en"], "name_kh": w["name_kh"], "dept": w["dept"], "title": w["title"], "entry": entry}
        pe = probation_end(entry)
        if first <= pe <= last:
            prob.append({**base, "probation_end": pe})
        for no in range(1, 60):
            s_, e_ = period(entry, no)
            if s_ > last:
                break
            c = mine.get(no)
            if c:
                s_, e_ = datetime.strptime(c["start"], "%Y-%m-%d").date(), datetime.strptime(c["end"], "%Y-%m-%d").date()
            if first <= s_ <= last:
                make.append({**base, "no": no, "start": s_, "end": e_, "made_on": (c or {}).get("made_on", ""), "signed_on": (c or {}).get("signed_on", ""),
                             "returned_on": (c or {}).get("returned_on", "")})
            if first <= e_ <= last:
                settle.append({**base, "no": no, "start": s_, "end": e_, "settled_on": (c or {}).get("settled_on", "")})
    return make, settle, prob


# ---------------------------------------------------------------- holiday work already paid on the day (the list you exclude from payroll)
def hp_rate(holiday_name):
    """Dollars per hour: a holiday = 2, a normal day = 1 (you can change both in the settings table)."""
    return get_setting("hp_rate_normal") if str(holiday_name or "").strip().upper() == "NORMAL DAY" else get_setting("hp_rate_holiday")


def hp_list(month=None):
    if month:
        return _rows("SELECT * FROM holiday_paid WHERE substr(day,1,7)=? ORDER BY day, worker_id", (month,))
    return _rows("SELECT * FROM holiday_paid ORDER BY day DESC, worker_id")


def hp_months():
    return [r["m"] for r in _rows("SELECT DISTINCT substr(day,1,7) AS m FROM holiday_paid ORDER BY m DESC")]


def hp_save(r):
    """r = dict(worker_id, name, dept, position, day, holiday_name, hours, amount, date_paid, method, exclude). Same worker + same day = replaced. Returns '' or an error text."""
    wid = str(r.get("worker_id", "")).strip().upper().replace(" ", "")
    if not ID_RE.match(wid):
        return f"'{wid}': the worker ID must look like DMO-10002."
    try:
        day = str(r["day"])[:10]
        datetime.strptime(day, "%Y-%m-%d")
    except Exception:
        return f"{wid}: the date is not valid."
    try:
        hours = float(r.get("hours") or 0)
        amount = float(r.get("amount") or 0)
    except (TypeError, ValueError):
        return f"{wid}: hours and amount must be numbers."
    if not (0 < hours <= 24):
        return f"{wid}: hours must be more than 0 and at most 24."
    if amount < 0 or amount > 10000:
        return f"{wid}: the amount does not look right."
    dp = str(r.get("date_paid") or "")[:10]
    if dp:
        try:
            datetime.strptime(dp, "%Y-%m-%d")
        except ValueError:
            return f"{wid}: the 'date paid' is not a date."
    ex = "YES" if str(r.get("exclude", "YES")).strip().upper() in ("YES", "Y", "TRUE", "1") else "NO"
    clean = lambda v, n=80: re.sub(r"\s+", " ", str(v or "")).strip()[:n]
    if str(clean(r.get("name"))).startswith(("=", "+", "-", "@")):
        return f"{wid}: the name must not start with = + - or @."
    con = connect()
    with con:
        con.execute("INSERT INTO holiday_paid(worker_id, name, dept, position, day, holiday_name, hours, days, amount, date_paid, method, exclude, created_at) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(worker_id, day) DO UPDATE SET name=excluded.name, dept=excluded.dept, position=excluded.position, "
                    "holiday_name=excluded.holiday_name, hours=excluded.hours, days=excluded.days, amount=excluded.amount, date_paid=excluded.date_paid, "
                    "method=excluded.method, exclude=excluded.exclude",
                    (wid, clean(r.get("name")), clean(r.get("dept"), 30).upper(), clean(r.get("position")), day, clean(r.get("holiday_name"), 40),
                     hours, round(hours / 8, 2), amount, dp, clean(r.get("method"), 30), ex, _now()))
    con.close()
    return ""


def hp_delete(ids):
    con = connect()
    with con:
        for i in ids:
            con.execute("DELETE FROM holiday_paid WHERE id=?", (int(i),))
    con.close()


def hp_import_xlsx(data):
    """Reads your 'Holiday Work List' Excel (sheet 'Holiday Work List'). Returns (added, errors list). The file is only read."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    except Exception as e:
        return 0, [f"I could not open the Excel file ({type(e).__name__})."]
    name = "Holiday Work List" if "Holiday Work List" in wb.sheetnames else wb.sheetnames[0]
    rows = list(wb[name].iter_rows(values_only=True))
    wb.close()
    head = next((i for i, r in enumerate(rows[:15]) if r and any(str(c or "").strip().lower() == "employee id" for c in r)), None)
    if head is None:
        return 0, ["I could not find the heading row ('Employee ID') in the first 15 lines."]
    idx = {str(c or "").strip().lower(): j for j, c in enumerate(rows[head])}
    need = ["employee id", "holiday date", "hours worked"]
    if any(k not in idx for k in need):
        return 0, ["The heading row needs the columns: Employee ID, Holiday Date, Hours Worked."]
    pos = {w["id"]: w for w in list_workers()}

    def cell(r, k):
        j = idx.get(k)
        return r[j] if j is not None and j < len(r) else None
    def as_day(v):
        if isinstance(v, datetime):
            return v.date().isoformat()
        if isinstance(v, date):
            return v.isoformat()
        d = _d_any(v)
        return d.isoformat() if d else ""
    added, errs = 0, []
    for n, r in enumerate(rows[head + 1:], start=head + 2):
        wid = str(cell(r, "employee id") or "").strip().upper()
        if not wid.startswith("DMO"):
            continue
        hol = str(cell(r, "holiday name") or "").strip()
        hours = cell(r, "hours worked")
        paid = cell(r, "amount already paid")
        rec = {"worker_id": wid, "name": cell(r, "employee name"), "dept": cell(r, "department"),
               "position": cell(r, "position") or (pos.get(wid) or {}).get("title", ""), "day": as_day(cell(r, "holiday date")),
               "holiday_name": hol, "hours": hours, "amount": paid if paid not in (None, "") else (float(hours or 0) * hp_rate(hol)),
               "date_paid": as_day(cell(r, "date paid")), "method": cell(r, "payment ref / method"), "exclude": cell(r, "exclude from oct payroll?") or "YES"}
        err = hp_save(rec)
        if err:
            errs.append(f"Line {n}: {err}")
        else:
            added += 1
    return added, errs


def _d_any(v):
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(v).strip()[:10], f).date()
        except ValueError:
            pass
    return None


def hp_export_xlsx(month):
    """Builds the same Excel you use: sheet 1 = one row per worker per day, sheet 2 = one row per worker with the hours under each date. Returns bytes."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    rows = hp_list(month)
    y, m = int(month[:4]), int(month[5:7])
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Holiday Work List"
    bold, fill = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    ws["A1"] = "Staff Who Worked on Public Holidays (Already Paid) - Exclude from Payroll"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"], ws["C2"] = "Prepared by:", ""
    ws["A3"], ws["C3"] = "Payroll month:", date(y, m, 1).strftime("%B %Y")
    heads = ["No.", "Employee ID", "Employee Name", "Department", "Position", "Holiday Date", "Holiday Name", "Day", "Hours Worked", "Days Worked",
             "Amount Already Paid", "Date Paid", "Payment Ref / Method", "Exclude from Payroll?"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.fill = bold, fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, r in enumerate(rows, 1):
        d = datetime.strptime(r["day"], "%Y-%m-%d").date()
        vals = [i, r["worker_id"], r["name"], r["dept"], r["position"], d, r["holiday_name"], d.strftime("%a"), r["hours"], r["days"],
                r["amount"], _d_any(r["date_paid"]) if r["date_paid"] else None, r["method"], r["exclude"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=4 + i, column=j, value=v)
            if j in (6, 12) and v:
                c.number_format = "DD-MMM-YYYY"
    last = 4 + len(rows)
    ws.cell(row=last + 2, column=3, value="TOTAL").font = bold
    ws.cell(row=last + 2, column=9, value=sum(r["hours"] for r in rows)).font = bold
    ws.cell(row=last + 2, column=10, value=sum(r["days"] for r in rows)).font = bold
    ws.cell(row=last + 2, column=11, value=sum(r["amount"] for r in rows)).font = bold
    ws.cell(row=last + 3, column=3, value="Total to exclude from payroll").font = bold
    ws.cell(row=last + 3, column=11, value=sum(r["amount"] for r in rows if r["exclude"] == "YES")).font = bold
    for j, w_ in enumerate([6, 13, 26, 14, 18, 14, 22, 6, 12, 11, 14, 14, 18, 14], 1):
        ws.column_dimensions[get_column_letter(j)].width = w_
    ws.freeze_panes = "A5"
    s2 = wb.create_sheet("Summary by Employee")
    s2["A1"] = "Holiday Work Summary - One Row per Employee (hours under each date)"
    s2["A1"].font = Font(bold=True, size=13)
    days = sorted({r["day"] for r in rows})
    heads2 = ["No.", "Employee ID", "Employee Name", "Department"] + [datetime.strptime(d, "%Y-%m-%d").date() for d in days] + ["Total Days", "Total Hours", "Amount Paid"]
    for j, h in enumerate(heads2, 1):
        c = s2.cell(row=3, column=j, value=h)
        c.font, c.fill = bold, fill
        if isinstance(h, date):
            c.number_format = "DD-MMM"
    people = {}
    for r in rows:
        p = people.setdefault(r["worker_id"], {"name": r["name"], "dept": r["dept"], "by": {}, "days": 0.0, "hours": 0.0, "amount": 0.0})
        p["by"][r["day"]] = p["by"].get(r["day"], 0) + r["hours"]
        p["days"] += r["days"]; p["hours"] += r["hours"]; p["amount"] += r["amount"]
    for i, (wid, p) in enumerate(sorted(people.items()), 1):
        vals = [i, wid, p["name"], p["dept"]] + [p["by"].get(d) for d in days] + [p["days"], p["hours"], p["amount"]]
        for j, v in enumerate(vals, 1):
            s2.cell(row=3 + i, column=j, value=v)
    t = 4 + len(people)
    s2.cell(row=t, column=3, value="All staff / hours / paid").font = bold
    s2.cell(row=t, column=4 + len(days) + 1, value=sum(p["days"] for p in people.values())).font = bold
    s2.cell(row=t, column=4 + len(days) + 2, value=sum(p["hours"] for p in people.values())).font = bold
    s2.cell(row=t, column=4 + len(days) + 3, value=sum(p["amount"] for p in people.values())).font = bold
    for j in range(1, len(heads2) + 1):
        s2.column_dimensions[get_column_letter(j)].width = 14 if j != 3 else 26
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


# ================================================================== stage 3+: recruitment, resigned, planned headcount, reminders, reports
SETTING_DEFAULTS = {"fdc_months": "24", "fdc_warn_days": "60", "unpaid_red_days": "7", "track_resign_days": "45", "absent_days": "3",
                    "hp_rate_holiday": "2", "hp_rate_normal": "1", "contract_salary": "210", "contract_allowance": "0"}
CANDIDATE_STATUS = ["Found", "Interview", "Hired", "Not coming", "Rejected"]
WARNING_KINDS = ["Verbal warning", "Written warning 1", "Written warning 2", "Final warning", "Other"]
REMINDER_KINDS = ["Reminder", "Maternity", "Other"]
FACTORY_KINDS = ["License / permit", "Insurance", "Fire safety", "Certificate", "Contract / agreement", "Other"]
FACTORY_DIR = os.path.join("web_work", "factory_docs")


def get_setting(key):
    r = _rows("SELECT value FROM settings WHERE key=?", (key,))
    return int(r[0]["value"]) if r and str(r[0]["value"]).isdigit() else int(SETTING_DEFAULTS[key])


def set_setting(key, value):
    if key not in SETTING_DEFAULTS:
        return
    con = connect()
    with con:
        con.execute("INSERT OR REPLACE INTO settings(key, value) VALUES(?,?)", (key, str(max(0, int(value)))))
    con.close()


def _iso_ok(v):
    if not v:
        return True
    try:
        datetime.strptime(str(v), "%Y-%m-%d")
        return True
    except ValueError:
        return False


def _d(v):
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except Exception:
        return None


# ---------------------------------------------------------------- recruitment
def next_worker_id():
    nums = [int(r["id"][4:]) for r in _rows("SELECT id FROM workers") if ID_RE.match(r["id"])]
    return "DMO-%05d" % ((max(nums) + 1) if nums else 10001)


def save_request(r):
    """r = dict(id optional, dept, position, need, asked_by, asked_on, deadline, note, closed). Returns '' or an error text."""
    dept = re.sub(r"\s+", " ", str(r.get("dept", ""))).strip().upper()[:40]
    pos = re.sub(r"\s+", " ", str(r.get("position", ""))).strip()[:60]
    if not dept or not pos:
        return "Type the department and the position."
    try:
        need = int(r.get("need"))
    except (TypeError, ValueError):
        return "How many people: type a number."
    if need < 1 or need > 5000:
        return "How many people: type a number from 1 to 5000."
    for f in ("asked_on", "deadline"):
        if not _iso_ok(r.get(f, "")):
            return f"{f.replace('_', ' ')} is not a date."
    con = connect()
    with con:
        if r.get("id"):
            con.execute("UPDATE recruit_requests SET dept=?, position=?, need=?, asked_by=?, asked_on=?, deadline=?, note=?, closed=? WHERE id=?",
                        (dept, pos, need, str(r.get("asked_by", ""))[:60], r.get("asked_on", ""), r.get("deadline", ""),
                         str(r.get("note", ""))[:300], 1 if r.get("closed") else 0, int(r["id"])))
        else:
            con.execute("INSERT INTO recruit_requests(dept, position, need, asked_by, asked_on, deadline, note, closed, created_at) VALUES(?,?,?,?,?,?,?,0,?)",
                        (dept, pos, need, str(r.get("asked_by", ""))[:60], r.get("asked_on", ""), r.get("deadline", ""),
                         str(r.get("note", ""))[:300], _now()))
    con.close()
    return ""


def delete_request(rid):
    con = connect()
    with con:
        con.execute("DELETE FROM candidates WHERE request_id=? AND worker_id=''", (int(rid),))
        con.execute("UPDATE candidates SET request_id=0 WHERE request_id=?", (int(rid),))
        con.execute("DELETE FROM recruit_requests WHERE id=?", (int(rid),))
    con.close()


def list_requests(include_closed=True):
    """Requests with how many were hired (status Hired, or already added to the worker book)."""
    reqs = _rows("SELECT * FROM recruit_requests ORDER BY closed, deadline = '', deadline, id")
    cands = _rows("SELECT request_id, status, worker_id FROM candidates")
    today = date.today()
    for q in reqs:
        mine = [c for c in cands if c["request_id"] == q["id"]]
        q["found"] = len(mine)
        q["hired"] = sum(1 for c in mine if c["status"] == "Hired" or c["worker_id"])
        q["left"] = max(0, q["need"] - q["hired"])
        dl = _d(q["deadline"])
        q["days_left"] = (dl - today).days if dl else None
        q["open"] = not q["closed"] and q["left"] > 0
    return reqs if include_closed else [q for q in reqs if q["open"]]


def rehire_check(id_card, exclude_candidate=None):
    """Same ID card number already known? Returns a list of text lines (old worker records, other candidates)."""
    n = re.sub(r"\s+", "", str(id_card or "")).upper()
    if len(n) < 4:
        return []
    out = []
    for w in _rows("SELECT id, name_en, dept, entry_date, stop_date, stop_date_raw, id_card FROM workers"):
        if re.sub(r"\s+", "", w["id_card"] or "").upper() == n:
            left = (" - resigned " + w["stop_date"]) if w["stop_date"] else (" - resigned (date unclear)" if w["stop_date_raw"] else " - STILL WORKING")
            out.append(f"{w['id']} {w['name_en']} ({w['dept']}), started {w['entry_date'] or '?'}{left}")
    for c in _rows("SELECT id, name_en, name_kh, status, worker_id, id_card FROM candidates"):
        if exclude_candidate and c["id"] == exclude_candidate:
            continue
        if re.sub(r"\s+", "", c["id_card"] or "").upper() == n:
            out.append(f"Candidate {c['name_en'] or c['name_kh']} is already on a recruitment list (status {c['status']})")
    return out


def add_candidate(c):
    """c = dict(request_id, name_en, name_kh, phone, id_card, sex, found_on, note). Returns (error, warnings list)."""
    en = re.sub(r"\s+", " ", str(c.get("name_en", ""))).strip()[:60]
    kh = re.sub(r"\s+", " ", str(c.get("name_kh", ""))).strip()[:60]
    if not en and not kh:
        return "Type the name of the person.", []
    if c.get("sex", "") not in ("", "M", "F"):
        return "Sex must be M or F.", []
    if not _iso_ok(c.get("found_on", "")):
        return "Date found is not a date.", []
    warns = rehire_check(c.get("id_card", ""))
    con = connect()
    with con:
        con.execute("INSERT INTO candidates(request_id, name_en, name_kh, phone, id_card, sex, found_on, status, note, created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                    (int(c.get("request_id") or 0), en, kh, str(c.get("phone", ""))[:30], str(c.get("id_card", ""))[:30], c.get("sex", ""),
                     c.get("found_on", ""), "Found", str(c.get("note", ""))[:300], _now()))
    con.close()
    return "", warns


def list_candidates(request_id=None):
    sql = "SELECT c.*, r.position, r.dept FROM candidates c LEFT JOIN recruit_requests r ON r.id=c.request_id"
    args = ()
    if request_id is not None:
        sql += " WHERE c.request_id=?"
        args = (int(request_id),)
    return _rows(sql + " ORDER BY c.id DESC", args)


def update_candidates(items):
    """items = list of dict(id, status, interview_on, start_on, phone, note). Returns an error text or ''."""
    for x in items:
        if x["status"] not in CANDIDATE_STATUS:
            return "Unknown status."
        for f in ("interview_on", "start_on"):
            if not _iso_ok(x.get(f, "")):
                return f"{f.replace('_', ' ')} is not a date."
    con = connect()
    with con:
        for x in items:
            con.execute("UPDATE candidates SET status=?, interview_on=?, start_on=?, phone=?, note=? WHERE id=? AND worker_id=''",
                        (x["status"], x.get("interview_on", ""), x.get("start_on", ""), str(x.get("phone", ""))[:30], str(x.get("note", ""))[:300], int(x["id"])))
    con.close()
    return ""


def delete_candidate(cid):
    con = connect()
    with con:
        con.execute("DELETE FROM candidates WHERE id=? AND worker_id=''", (int(cid),))
    con.close()


def hire_to_worker(cand_id, wid, dept, title, entry_iso):
    """Makes the candidate a worker in the worker book (new ID) and adds contract 1 (nothing made yet). Returns '' or an error text."""
    c = _rows("SELECT * FROM candidates WHERE id=?", (int(cand_id),))
    if not c:
        return "I cannot find this person."
    c = c[0]
    if c["worker_id"]:
        return f"This person is already in the worker book as {c['worker_id']}."
    wid = str(wid).strip().upper()
    if not _iso_ok(entry_iso) or not entry_iso:
        return "Type the entry date."
    err = save_worker({"id": wid, "name_en": c["name_en"], "name_kh": c["name_kh"], "dept": dept, "title": title, "sex": c["sex"],
                       "entry_date": entry_iso, "id_card": c["id_card"], "phone": c["phone"], "remark": "Recruited", "fp_registered": "N"}, new=True)
    if err:
        return err
    con = connect()
    with con:
        con.execute("UPDATE candidates SET worker_id=?, status='Hired', start_on=? WHERE id=?", (wid, entry_iso, int(cand_id)))
    con.close()
    s, e = period(_d(entry_iso), 1)
    save_contract({"worker_id": wid, "no": 1, "start": s.isoformat(), "end": e.isoformat()})
    return ""


# ---------------------------------------------------------------- resigned workers
def resigned_rows(today=None):
    """Every resigned worker with the 4 steps. Workers resigned before the tracking window with no record are 'untracked'."""
    today = today or date.today()
    window = get_setting("track_resign_days")
    rec = {r["worker_id"]: r for r in _rows("SELECT * FROM resignations")}
    out = []
    for w in list_workers():
        if w["status"] != "Resigned":
            continue
        r = rec.get(w["id"])
        sd = _d(w["stop_date"])
        days = (today - sd).days if sd else None
        base = {"worker_id": w["id"], "name": w["name_en"], "name_kh": w["name_kh"], "dept": w["dept"], "title": w["title"], "nssf": w["nssf"],
                "stop_date": w["stop_date"], "stop_raw": w["stop_date_raw"], "days": days, "record": bool(r)}
        if r:
            base.update({k: r[k] for k in ("reason", "paper_done", "paper_on", "paid", "paid_on", "items_back", "items_on", "nssf_reported", "nssf_on", "note")})
        else:
            base.update({"reason": "", "paper_done": 0, "paper_on": "", "paid": 0, "paid_on": "", "items_back": 0, "items_on": "",
                         "nssf_reported": 0, "nssf_on": "", "note": ""})
        base["tracked"] = bool(r) or (days is not None and days <= window)
        base["open"] = base["tracked"] and not (base["paper_done"] and base["paid"] and base["items_back"] and base["nssf_reported"])
        out.append(base)
    out.sort(key=lambda x: (x["stop_date"] or "0000"), reverse=True)
    return out


def save_resign(x):
    """x = dict(worker_id, reason, paper_done, paper_on, paid, paid_on, items_back, items_on, nssf_reported, nssf_on, note)."""
    if not ID_RE.match(str(x.get("worker_id", ""))):
        return "bad worker id"
    for f in ("paper_on", "paid_on", "items_on", "nssf_on"):
        if not _iso_ok(x.get(f, "")):
            return f"{x['worker_id']}: {f.replace('_', ' ')} is not a date."
    flags = {}
    for flag, on in (("paper_done", "paper_on"), ("paid", "paid_on"), ("items_back", "items_on"), ("nssf_reported", "nssf_on")):
        flags[flag] = 1 if x.get(flag) else 0
        flags[on] = x.get(on, "") if flags[flag] else ""
    con = connect()
    with con:
        con.execute("INSERT OR REPLACE INTO resignations(worker_id, reason, paper_done, paper_on, paid, paid_on, items_back, items_on, nssf_reported, nssf_on, note) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (x["worker_id"], str(x.get("reason", ""))[:120], flags["paper_done"], flags["paper_on"], flags["paid"], flags["paid_on"],
                     flags["items_back"], flags["items_on"], flags["nssf_reported"], flags["nssf_on"], str(x.get("note", ""))[:300]))
    con.close()
    return ""


def mark_resigned(wid, stop_iso, reason=""):
    """A worker leaves: sets the stop date in the worker book and starts the resign tracker line."""
    w = get_worker(wid)
    if not w:
        return f"{wid} is not in the worker book."
    if not stop_iso or not _iso_ok(stop_iso):
        return "Type the last working day."
    con = connect()
    with con:
        con.execute("UPDATE workers SET stop_date=?, stop_date_raw='', updated_at=? WHERE id=?", (stop_iso, _now(), wid))
    con.close()
    return save_resign({"worker_id": wid, "reason": reason})


def mark_old_resigned_done(before_iso):
    """Older resigned workers that were never tracked: write a finished record so they stop showing. Returns how many."""
    n = 0
    for r in resigned_rows():
        if not r["record"] and r["stop_date"] and r["stop_date"] < before_iso:
            save_resign({"worker_id": r["worker_id"], "reason": "(old, before I started tracking)", "paper_done": 1, "paid": 1,
                         "items_back": 1, "nssf_reported": 1, "paper_on": r["stop_date"], "paid_on": r["stop_date"],
                         "items_on": r["stop_date"], "nssf_on": r["stop_date"]})
            n += 1
    return n


# ---------------------------------------------------------------- planned headcount + dashboard numbers
def get_plan():
    return {r["dept"]: r["planned"] for r in _rows("SELECT * FROM dept_plan")}


def save_plan(plan):
    con = connect()
    with con:
        for d, n in plan.items():
            d = str(d).strip().upper()
            if d:
                con.execute("INSERT OR REPLACE INTO dept_plan(dept, planned) VALUES(?,?)", (d, max(0, int(n or 0))))
    con.close()


def headcount_on(on, ws=None):
    """Workers who were working on date `on` (entry date empty counts as already working)."""
    ws = ws if ws is not None else list_workers()
    out = []
    for w in ws:
        ent = _d(w["entry_date"])
        stp = _d(w["stop_date"])
        if w["stop_date_raw"] and not stp:
            continue
        if ent and ent > on:
            continue
        if stp and stp <= on:
            continue
        out.append(w)
    return out


def dept_table(ws=None):
    """Rows: dept, planned, real, gap (planned - real), asked (open requests still needed)."""
    ws = ws if ws is not None else list_workers()
    plan = get_plan()
    real = {}
    for w in ws:
        if w["status"] == "Working":
            real[w["dept"] or "(no department)"] = real.get(w["dept"] or "(no department)", 0) + 1
    asked = {}
    for q in list_requests(include_closed=False):
        asked[q["dept"]] = asked.get(q["dept"], 0) + q["left"]
    names = sorted(set(plan) | set(real) | set(asked))
    return [{"dept": d, "planned": plan.get(d, 0), "real": real.get(d, 0), "gap": plan.get(d, 0) - real.get(d, 0), "asked": asked.get(d, 0)} for d in names]


def month_report(year, month):
    """MLVT-style monthly numbers: start, hired, resigned, end, women/men, by department."""
    import calendar
    first = date(year, month, 1)
    last = date(year, month, calendar.monthrange(year, month)[1])
    ws = list_workers()
    start = headcount_on(first - __import__("datetime").timedelta(days=1), ws)
    end = headcount_on(last, ws)
    hired = [w for w in ws if _d(w["entry_date"]) and first <= _d(w["entry_date"]) <= last]
    left = [w for w in ws if _d(w["stop_date"]) and first <= _d(w["stop_date"]) <= last]
    rows = {}
    for tag, lst in (("start", start), ("hired", hired), ("left", left), ("end", end)):
        for w in lst:
            d = rows.setdefault(w["dept"] or "(no department)", {"dept": w["dept"] or "(no department)", "start": 0, "hired": 0, "left": 0, "end": 0, "end_women": 0})
            d[tag] += 1
            if tag == "end" and w["sex"] == "F":
                d["end_women"] += 1
    no_entry = sum(1 for w in ws if not w["entry_date"])
    return {"first": first, "last": last, "start": len(start), "hired": len(hired), "left": len(left), "end": len(end),
            "women": sum(w["sex"] == "F" for w in end), "men": sum(w["sex"] == "M" for w in end),
            "rows": sorted(rows.values(), key=lambda r: r["dept"]), "hired_list": hired, "left_list": left, "no_entry": no_entry}


def shift_counts():
    """For every day you checked and saved: how many workers came on the day shift, the night shift (evening shift 16:00-00:00 counts as night) and the office.
    A worker counts when he has a first fingerprint scan on that work day. Returns a list of dicts, newest day first."""
    dept = {w["id"]: (w["dept"] or "").upper() for w in list_workers()}
    out = {}
    for r in _rows("SELECT day, worker_id, night, in_time FROM records WHERE in_time IS NOT NULL AND in_time != ''"):
        d = out.setdefault(r["day"], {"day": r["day"], "day_n": 0, "night_n": 0, "office_n": 0, "by_dept": {}})
        dp = dept.get(r["worker_id"], "") or "(not in worker book)"
        kind = "office_n" if dp == "OFFICE" else ("night_n" if r["night"] else "day_n")
        d[kind] += 1
        b = d["by_dept"].setdefault(dp, {"day_n": 0, "night_n": 0, "office_n": 0})
        b[kind] += 1
    rows = sorted(out.values(), key=lambda x: x["day"], reverse=True)
    for d in rows:
        d["total"] = d["day_n"] + d["night_n"] + d["office_n"]
    return rows


def monthly_moves(first_month=None, last_month=None):
    """Hired (entry date) and resigned (stop date) workers for every month, with the number working at the end of the month.
    Returns a list of dicts, oldest month first, and the two worker lists by month."""
    import calendar
    ws = list_workers()
    hired, left = {}, {}
    for w in ws:
        e, t = _d(w["entry_date"]), _d(w["stop_date"])
        if e:
            hired.setdefault(f"{e:%Y-%m}", []).append(w)
        if t:
            left.setdefault(f"{t:%Y-%m}", []).append(w)
    keys = sorted(set(hired) | set(left))
    if not keys:
        return [], hired, left
    a = first_month or keys[0]
    b = last_month or max(keys[-1], f"{date.today():%Y-%m}")
    y, m = int(a[:4]), int(a[5:7])
    rows = []
    while f"{y:04d}-{m:02d}" <= b:
        key = f"{y:04d}-{m:02d}"
        end = date(y, m, calendar.monthrange(y, m)[1])
        rows.append({"month": key, "hired": len(hired.get(key, [])), "left": len(left.get(key, [])),
                     "net": len(hired.get(key, [])) - len(left.get(key, [])), "working_end": len(headcount_on(end, ws))})
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return rows, hired, left


def save_snapshot(year, month):
    import json
    m = month_report(year, month)
    con = connect()
    with con:
        con.execute("INSERT OR REPLACE INTO snapshots(month, working, women, men, hired, resigned, depts, saved_at) VALUES(?,?,?,?,?,?,?,?)",
                    (f"{year:04d}-{month:02d}", m["end"], m["women"], m["men"], m["hired"], m["left"],
                     json.dumps({r["dept"]: r["end"] for r in m["rows"]}), _now()))
    con.close()


def list_snapshots():
    return _rows("SELECT * FROM snapshots ORDER BY month DESC")


# ---------------------------------------------------------------- reminders and warning letters
def add_reminder(due, kind, wid, text):
    if not _iso_ok(due) or not due:
        return "Type the date."
    text = re.sub(r"\s+", " ", str(text or "")).strip()[:300]
    if not text:
        return "Type what to remember."
    wid = str(wid or "").strip().upper()
    if wid and not ID_RE.match(wid):
        return "The worker ID must look like DMO-10002 (or leave it empty)."
    if kind not in REMINDER_KINDS:
        kind = "Reminder"
    con = connect()
    with con:
        con.execute("INSERT INTO reminders(due, kind, worker_id, text, done, created_at) VALUES(?,?,?,?,0,?)", (due, kind, wid, text, _now()))
    con.close()
    return ""


def list_reminders(only_open=True):
    return _rows("SELECT * FROM reminders" + (" WHERE done=0" if only_open else "") + " ORDER BY done, due, id")


def set_reminders_done(done_ids, delete_ids=()):
    con = connect()
    with con:
        for i in done_ids:
            con.execute("UPDATE reminders SET done=1 WHERE id=?", (int(i),))
        for i in delete_ids:
            con.execute("DELETE FROM reminders WHERE id=?", (int(i),))
    con.close()


def add_warning(wid, on_date, kind, reason):
    if not ID_RE.match(str(wid)):
        return "bad worker id"
    if not on_date or not _iso_ok(on_date):
        return "Type the date of the letter."
    if kind not in WARNING_KINDS:
        return "Choose the kind."
    con = connect()
    with con:
        con.execute("INSERT INTO warnings(worker_id, on_date, kind, reason, created_at) VALUES(?,?,?,?,?)",
                    (wid, on_date, kind, str(reason or "")[:300], _now()))
    con.close()
    return ""


def list_warnings(wid=None):
    if wid:
        return _rows("SELECT * FROM warnings WHERE worker_id=? ORDER BY on_date DESC, id DESC", (wid,))
    return _rows("SELECT w.*, k.name_en, k.dept FROM warnings w LEFT JOIN workers k ON k.id=w.worker_id ORDER BY on_date DESC, w.id DESC")


def delete_warning(i):
    con = connect()
    with con:
        con.execute("DELETE FROM warnings WHERE id=?", (int(i),))
    con.close()


# ---------------------------------------------------------------- factory documents
def add_factory_doc(name, kind, expiry, note, filename=None, data=None):
    name = re.sub(r"\s+", " ", str(name or "")).strip()[:80]
    if not name:
        return "Type the name of the document."
    if expiry and not _iso_ok(expiry):
        return "Expiry is not a date."
    saved = ""
    if filename and data:
        ext = os.path.splitext(filename)[1].lower()
        if ext not in ALLOWED_FILES:
            return f"File type {ext or '(none)'} is not allowed. Use: " + ", ".join(ALLOWED_FILES)
        if len(data) > MAX_FILE_MB * 1024 * 1024:
            return f"The file is bigger than {MAX_FILE_MB} MB."
        os.makedirs(FACTORY_DIR, exist_ok=True)
        saved = f"{datetime.now():%Y%m%d_%H%M%S}__{safe_file_name(os.path.splitext(filename)[0])}{ext}"
        with open(os.path.join(FACTORY_DIR, saved), "wb") as f:
            f.write(data)
    con = connect()
    with con:
        con.execute("INSERT INTO factory_docs(name, kind, expiry, note, file, created_at) VALUES(?,?,?,?,?,?)",
                    (name, kind if kind in FACTORY_KINDS else "Other", expiry or "", str(note or "")[:300], saved, _now()))
    con.close()
    return ""


def list_factory_docs():
    rows = _rows("SELECT * FROM factory_docs ORDER BY expiry = '', expiry, id")
    today = date.today()
    for r in rows:
        e = _d(r["expiry"])
        r["days"] = (e - today).days if e else None
        p = os.path.join(FACTORY_DIR, os.path.basename(r["file"])) if r["file"] else ""
        r["path"] = p if p and os.path.isfile(p) else ""
    return rows


def delete_factory_doc(i):
    con = connect()
    with con:
        r = con.execute("SELECT file FROM factory_docs WHERE id=?", (int(i),)).fetchone()
        con.execute("DELETE FROM factory_docs WHERE id=?", (int(i),))
    con.close()
    if r and r[0]:
        try:
            os.remove(os.path.join(FACTORY_DIR, os.path.basename(r[0])))
        except OSError:
            pass


# ---------------------------------------------------------------- Today page: alerts
def absent_streaks(min_days=None):
    """Workers whose last checked days were all code A (absent). Uses the days you saved in Daily work (decisions table)."""
    min_days = min_days or get_setting("absent_days")
    days = [d["day"] for d in _adb.days_list()]
    if not days:
        return []
    codes = {d: {x["worker_id"]: (x["code"] or "").upper() for x in _adb.day_decisions(d)} for d in days[:30]}
    out = []
    for w in list_workers():
        if w["status"] != "Working":
            continue
        n = 0
        for d in days[:30]:
            if codes[d].get(w["id"]) == "A":
                n += 1
            else:
                break
        if n >= min_days:
            out.append({"worker_id": w["id"], "name": w["name_en"], "dept": w["dept"], "days_in_a_row": n, "last_day": days[0]})
    return out


def codes_in_month(year, month):
    """Count of each code per worker for a month, from your saved decisions."""
    prefix = f"{year:04d}-{month:02d}-"
    names = {w["id"]: (w["name_en"], w["dept"]) for w in list_workers()}
    table = {}
    for r in _rows("SELECT worker_id, code FROM decisions WHERE day LIKE ?", (prefix + "%",)):
        code = (r["code"] or "").upper()
        if not code:
            continue
        t = table.setdefault(r["worker_id"], {})
        t[code] = t.get(code, 0) + 1
    return table, names


def birthdays(within_days=7, today=None):
    today = today or date.today()
    out = []
    for w in list_workers():
        b = _d(w["birthday"])
        if w["status"] != "Working" or not b:
            continue
        try:
            nxt = b.replace(year=today.year)
        except ValueError:
            nxt = date(today.year, 3, 1)
        if nxt < today:
            try:
                nxt = b.replace(year=today.year + 1)
            except ValueError:
                nxt = date(today.year + 1, 3, 1)
        d = (nxt - today).days
        if d <= within_days:
            out.append({"worker_id": w["id"], "name": w["name_en"], "dept": w["dept"], "birthday": nxt.isoformat(), "in_days": d, "turns": nxt.year - b.year})
    return sorted(out, key=lambda r: r["in_days"])


def fdc_rows(today=None):
    """Workers close to (or past) the limit of fixed-duration contracts counted from the entry date. The limit is a setting you can change."""
    today = today or date.today()
    lim, warn = get_setting("fdc_months"), get_setting("fdc_warn_days")
    out = []
    for w in list_workers():
        e = _d(w["entry_date"])
        if w["status"] != "Working" or not e:
            continue
        end = add_months(e, lim)
        d = (end - today).days
        if d <= warn:
            out.append({"worker_id": w["id"], "name": w["name_en"], "dept": w["dept"], "entry": w["entry_date"], "limit_date": end.isoformat(), "days": d})
    return sorted(out, key=lambda r: r["days"])


def alerts(today=None):
    """All things that need you today. Each: dict(level red/amber/green, title, rows list of dict, where)."""
    today = today or date.today()
    A = []

    def add(level, title, rows, where=""):
        if rows:
            A.append({"level": level, "title": title, "rows": rows, "where": where})

    ws = list_workers()
    # contracts
    cr = contract_rows(today, 14)
    add("red", "Contract is running but was never made", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Contract": r["no"], "Ended / ends": r["end"].isoformat()} for r in cr if r["state"] == "late"], "Contracts")
    cr30 = contract_rows(today, 30)
    add("red", "New contract starts within 14 days and is not made", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Contract": r["no"], "Starts": r["start"].isoformat(), "In days": r["days_to_start"]} for r in cr30 if r["state"] == "due" and r["days_to_start"] <= 14], "Contracts")
    add("amber", "New contract starts in 15 to 30 days: start making it", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Contract": r["no"], "Starts": r["start"].isoformat(), "In days": r["days_to_start"]} for r in cr30 if r["state"] == "due" and r["days_to_start"] > 14], "Contracts")
    sr = settlement_rows(today, 30)
    add("red", "Contract ended: the 5% end-of-contract money (结算) is not marked as paid", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Contract": r["no"], "Ended": r["end"].isoformat()} for r in sr if r["days_left"] < 0], "Contracts")
    add("amber", "Contract ends within 30 days: prepare the 5% money (结算)", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Contract": r["no"], "Ends": r["end"].isoformat(), "In days": r["days_left"]} for r in sr if r["days_left"] >= 0], "Contracts")
    add("amber", "Contract made, not signed yet", [{"Worker": r["worker_id"], "Name": r["name"], "Contract": r["no"], "Made on": r["made_on"]} for r in cr if r["state"] == "open" and not r["signed_on"]], "Contracts")
    add("amber", "Contract signed, not given back to the worker", [{"Worker": r["worker_id"], "Name": r["name"], "Contract": r["no"], "Signed on": r["signed_on"]} for r in cr if r["state"] == "open" and r["signed_on"] and not r["returned_on"]], "Contracts")
    # resigned
    red_days = get_setting("unpaid_red_days")
    rr = [r for r in resigned_rows(today) if r["open"]]
    add("red", f"Resigned and NOT paid for more than {red_days} days", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Left on": r["stop_date"], "Days": r["days"]} for r in rr if not r["paid"] and r["days"] is not None and r["days"] > red_days], "Resigned")
    add("amber", "Resigned, not paid yet", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Left on": r["stop_date"], "Days": r["days"]} for r in rr if not r["paid"] and not (r["days"] is not None and r["days"] > red_days)], "Resigned")
    add("amber", "Resigned: resign paper not filled", [{"Worker": r["worker_id"], "Name": r["name"], "Left on": r["stop_date"]} for r in rr if not r["paper_done"]], "Resigned")
    add("amber", "Resigned: items (ID card, uniform) not returned", [{"Worker": r["worker_id"], "Name": r["name"], "Left on": r["stop_date"]} for r in rr if not r["items_back"]], "Resigned")
    add("amber", "Resigned: NSSF stop not reported", [{"Worker": r["worker_id"], "Name": r["name"], "NSSF": r["nssf"], "Left on": r["stop_date"]} for r in rr if not r["nssf_reported"]], "Resigned")
    # recruitment
    rq = list_requests(include_closed=False)
    add("red", "Recruitment deadline passed and people are still missing", [{"Position": q["position"], "Dept": q["dept"], "Asked by": q["asked_by"], "Still need": q["left"], "Deadline": q["deadline"]} for q in rq if q["days_left"] is not None and q["days_left"] < 0], "Recruitment")
    add("amber", "Recruitment deadline in the next 7 days", [{"Position": q["position"], "Dept": q["dept"], "Still need": q["left"], "Deadline": q["deadline"], "In days": q["days_left"]} for q in rq if q["days_left"] is not None and 0 <= q["days_left"] <= 7], "Recruitment")
    cands = list_candidates()
    add("amber", "Interview today or soon, no answer yet", [{"Name": c["name_en"] or c["name_kh"], "Phone": c["phone"], "Interview": c["interview_on"], "For": c["position"]} for c in cands if c["status"] == "Interview" and c["interview_on"] and _d(c["interview_on"]) and _d(c["interview_on"]) <= today + __import__("datetime").timedelta(days=2) and not c["worker_id"]], "Recruitment")
    add("amber", "Hired, ready to put in the worker book", [{"Name": c["name_en"] or c["name_kh"], "Phone": c["phone"], "Start on": c["start_on"], "For": c["position"]} for c in cands if c["status"] == "Hired" and not c["worker_id"]], "Recruitment")
    # attendance
    add("red", f"Absent {get_setting('absent_days')} or more days in a row", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Days in a row": r["days_in_a_row"]} for r in absent_streaks()], "Daily work")
    try:
        lets = _adb.letters_list()
    except Exception:
        lets = []
    add("amber", "Missed-punch letter: worker not told yet", [{"Worker": x["worker_id"], "Name": x["name"], "Day": x["day"]} for x in lets if not x["told_on"] and not x["returned_on"]], "Letters")
    add("amber", "Missed-punch letter: told, still waiting", [{"Worker": x["worker_id"], "Name": x["name"], "Day": x["day"], "Told on": x["told_on"]} for x in lets if x["told_on"] and not x["returned_on"]], "Letters")
    # workers
    working = [w for w in ws if w["status"] == "Working"]
    add("amber", "Working without ID card copy or other required documents", [{"Worker": w["id"], "Name": w["name_en"], "Dept": w["dept"], "Documents missing": w["docs_missing"]} for w in working if w["docs_missing"]][:300], "Workers")
    add("amber", "Close to the fixed-contract limit (" + str(get_setting("fdc_months")) + " months)", [{"Worker": r["worker_id"], "Name": r["name"], "Entry": r["entry"], "Limit date": r["limit_date"], "Days": r["days"]} for r in fdc_rows(today)], "Workers")
    add("amber", "Working without an entry date", [{"Worker": w["id"], "Name": w["name_en"]} for w in working if not w["entry_date"]], "Workers")
    # reminders, factory documents
    rem = [r for r in list_reminders() if _d(r["due"]) and _d(r["due"]) <= today + __import__("datetime").timedelta(days=7)]
    add("red", "Reminders overdue", [{"Date": r["due"], "Kind": r["kind"], "Worker": r["worker_id"], "What": r["text"]} for r in rem if _d(r["due"]) < today], "Reminders")
    add("amber", "Reminders in the next 7 days", [{"Date": r["due"], "Kind": r["kind"], "Worker": r["worker_id"], "What": r["text"]} for r in rem if _d(r["due"]) >= today], "Reminders")
    fd = list_factory_docs()
    add("red", "Factory document expired", [{"Document": f["name"], "Expired": f["expiry"]} for f in fd if f["days"] is not None and f["days"] < 0], "Factory documents")
    add("amber", "Factory document expires in 30 days", [{"Document": f["name"], "Expires": f["expiry"], "In days": f["days"]} for f in fd if f["days"] is not None and 0 <= f["days"] <= 30], "Factory documents")
    # month end
    import calendar
    last = calendar.monthrange(today.year, today.month)[1]
    if today.day >= last - 2 and not any(s["month"] == f"{today.year:04d}-{today.month:02d}" for s in list_snapshots()):
        add("amber", "Month is ending: save the monthly headcount", [{"Month": f"{today:%B %Y}"}], "Reports")
    # good news
    add("green", "Birthdays in the next 7 days", [{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Birthday": r["birthday"], "In days": r["in_days"], "Turns": r["turns"]} for r in birthdays()], "Workers")
    order = {"red": 0, "amber": 1, "green": 2}
    A.sort(key=lambda a: order[a["level"]])
    return A


def boss_summary(today=None):
    """Plain text for the weekly summary to the boss."""
    import datetime as _dt
    today = today or date.today()
    mon = today - _dt.timedelta(days=today.weekday())
    sun = mon + _dt.timedelta(days=6)
    ws = list_workers()
    working = [w for w in ws if w["status"] == "Working"]
    hired = [w for w in ws if _d(w["entry_date"]) and mon <= _d(w["entry_date"]) <= sun]
    left = [w for w in ws if _d(w["stop_date"]) and mon <= _d(w["stop_date"]) <= sun]
    L = [f"HR weekly summary: {mon:%d %b} to {sun:%d %b %Y}", "",
         f"Workers now: {len(working)} (women {sum(w['sex'] == 'F' for w in working)}, men {sum(w['sex'] == 'M' for w in working)})",
         f"Started this week: {len(hired)}   Resigned this week: {len(left)}", ""]
    dt = [r for r in dept_table(ws) if r["planned"] or r["asked"]]
    if dt:
        L.append("Departments (planned / real / to recruit):")
        for r in dt:
            L.append(f"  {r['dept']}: {r['planned']} / {r['real']} / {max(r['gap'], 0)}")
        L.append("")
    rq = list_requests(include_closed=False)
    if rq:
        L.append("Recruitment requests still open:")
        for q in rq:
            L.append(f"  {q['position']} ({q['dept']}), asked by {q['asked_by'] or '?'}: hired {q['hired']} of {q['need']}, deadline {q['deadline'] or 'none'}")
        L.append("")
    cr = contract_rows(today, 30)
    L.append(f"Contracts: {sum(r['state'] == 'due' for r in cr)} starting within 30 days and not made; {sum(r['state'] == 'late' for r in cr)} running without a made contract.")
    rr = [r for r in resigned_rows(today) if r["open"]]
    L.append(f"Resigned: {sum(not r['paid'] for r in rr)} not paid yet, {sum(not r['paper_done'] for r in rr)} resign paper not filled, {sum(not r['nssf_reported'] for r in rr)} NSSF stop not reported.")
    L.append(f"Absent {get_setting('absent_days')}+ days in a row: {len(absent_streaks())}")
    return "\n".join(L)


def no_fingerprint_ids():
    """Working workers whose fingerprint is not registered yet (you tick 'registered' in the profile)."""
    return {w["id"] for w in list_workers() if w["status"] == "Working" and w["fp_registered"] == "N"}
