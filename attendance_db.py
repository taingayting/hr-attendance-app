"""
The attendance database (SQLite): ONE file on your laptop, memory/attendance.db. Free, no server, no account.
It remembers every day you check: the paper list, what the fingerprint showed, the remarks, your reasons and corrections,
so you can look up any worker or any day later.  To back it up, press the backup button in the website (or copy the file).
"""
import os, re, shutil, sqlite3, glob
from datetime import datetime

DB = "memory/attendance.db"
ARCHIVE = os.path.join("web_work", "archive")


def connect():
    os.makedirs("memory", exist_ok=True)
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    con.executescript("""
    CREATE TABLE IF NOT EXISTS days(
        day TEXT PRIMARY KEY, checked_at TEXT, filled_at TEXT, cells_written INTEGER DEFAULT 0,
        n_workers INTEGER, n_same INTEGER, n_attention INTEGER,
        paper_file TEXT, photo_file TEXT, fingerprint_file TEXT, attendance_file TEXT);
    CREATE TABLE IF NOT EXISTS records(
        day TEXT, worker_id TEXT, name TEXT, paper_hours TEXT, paper_ot TEXT, in_time TEXT, out_time TEXT,
        finger_ot TEXT, punches TEXT, status TEXT, remark TEXT, ot_filled TEXT, night INTEGER,
        PRIMARY KEY(day, worker_id));
    CREATE TABLE IF NOT EXISTS reasons(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT, worker_id TEXT, reason TEXT, final_ot TEXT, saved_at TEXT);
    CREATE TABLE IF NOT EXISTS corrections(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT, worker_id TEXT, field TEXT, was TEXT, now TEXT, saved_at TEXT);
    CREATE TABLE IF NOT EXISTS decisions(
        day TEXT, worker_id TEXT, hours TEXT, ot TEXT, code TEXT, reason TEXT, saved_at TEXT,
        PRIMARY KEY(day, worker_id));
    CREATE TABLE IF NOT EXISTS letters(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT, worker_id TEXT, name TEXT, dept TEXT, problem TEXT, punches TEXT,
        told_on TEXT DEFAULT '', returned_on TEXT DEFAULT '', note TEXT DEFAULT '', created_at TEXT,
        UNIQUE(day, worker_id));
    CREATE TABLE IF NOT EXISTS fp_outages(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT, from_time TEXT DEFAULT '', to_time TEXT DEFAULT '', dept TEXT DEFAULT '',
        note TEXT DEFAULT '', created_at TEXT);
    CREATE TABLE IF NOT EXISTS worker_shift(worker_id TEXT PRIMARY KEY, shift TEXT, updated TEXT);
    CREATE TABLE IF NOT EXISTS resigned_workers(worker_id TEXT PRIMARY KEY, since_day TEXT, noted_at TEXT);
    CREATE TABLE IF NOT EXISTS worker_times(worker_id TEXT PRIMARY KEY, t_start TEXT, t_end TEXT, updated TEXT);
    CREATE TABLE IF NOT EXISTS leave_notes(
        id INTEGER PRIMARY KEY AUTOINCREMENT, worker_id TEXT, day TEXT, kind TEXT, at_time TEXT DEFAULT '', note TEXT DEFAULT '',
        photo TEXT DEFAULT '', created_at TEXT);
    CREATE TABLE IF NOT EXISTS ot_questions(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT, worker_id TEXT, name TEXT, dept TEXT, shift TEXT DEFAULT '', paper_ot TEXT DEFAULT '', finger_ot TEXT DEFAULT '',
        asked_on TEXT DEFAULT '', answer TEXT DEFAULT '', note TEXT DEFAULT '', created_at TEXT, UNIQUE(day, worker_id));
    CREATE TABLE IF NOT EXISTS worker_cases(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT, worker_id TEXT, kind TEXT, name TEXT DEFAULT '', dept TEXT DEFAULT '', shift TEXT DEFAULT '',
        paper_ot TEXT DEFAULT '', finger_ot TEXT DEFAULT '', detail TEXT DEFAULT '', asked_on TEXT DEFAULT '', answer TEXT DEFAULT '', note TEXT DEFAULT '',
        created_at TEXT, UNIQUE(day, worker_id, kind));
    CREATE TABLE IF NOT EXISTS step3_draft(
        day TEXT, worker_id TEXT, hours TEXT DEFAULT '', ot TEXT DEFAULT '', code TEXT DEFAULT '', reason TEXT DEFAULT '', le TEXT DEFAULT '', letter INTEGER DEFAULT 0,
        saved_at TEXT, PRIMARY KEY(day, worker_id));
    CREATE TABLE IF NOT EXISTS notes(
        id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT, worker_id TEXT DEFAULT '', text TEXT, due TEXT DEFAULT '',
        snooze_until TEXT DEFAULT '', done_at TEXT DEFAULT '');
    CREATE TABLE IF NOT EXISTS list_notes(
        id INTEGER PRIMARY KEY AUTOINCREMENT, worker_id TEXT, day TEXT, text TEXT, created_at TEXT);
    CREATE TABLE IF NOT EXISTS pages(
        id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT DEFAULT '', body TEXT DEFAULT '', folder TEXT DEFAULT '', tags TEXT DEFAULT '',
        worker_id TEXT DEFAULT '', pinned INTEGER DEFAULT 0, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS page_files(
        id INTEGER PRIMARY KEY AUTOINCREMENT, page_id INTEGER, name TEXT, path TEXT, created_at TEXT);
    CREATE INDEX IF NOT EXISTS i_rec_worker ON records(worker_id);
    CREATE INDEX IF NOT EXISTS i_reason_worker ON reasons(worker_id);
    """)
    ocols = {r[1] for r in con.execute("PRAGMA table_info(fp_outages)").fetchall()}
    if "to_day" not in ocols:
        con.execute("ALTER TABLE fp_outages ADD COLUMN to_day TEXT DEFAULT ''")
    lcols = {r[1] for r in con.execute("PRAGMA table_info(leave_notes)").fetchall()}
    for c in ("to_day", "code", "medical"):
        if c not in lcols:
            con.execute(f"ALTER TABLE leave_notes ADD COLUMN {c} TEXT DEFAULT ''")
    if not con.execute("SELECT 1 FROM worker_cases LIMIT 1").fetchone():      # the OT cases of v4.3 move to the new list
        con.execute("""INSERT OR IGNORE INTO worker_cases(day, worker_id, kind, name, dept, shift, paper_ot, finger_ot, detail, asked_on, answer, note, created_at)
                       SELECT day, worker_id, 'OT different', name, dept, shift, paper_ot, finger_ot, 'paper ' || paper_ot || 'H, fingerprint ' || finger_ot || 'H',
                              asked_on, answer, note, created_at FROM ot_questions""")
    dcols = {r[1] for r in con.execute("PRAGMA table_info(decisions)").fetchall()}
    if "le" not in dcols:
        con.execute("ALTER TABLE decisions ADD COLUMN le TEXT DEFAULT ''")
    cols = {r[1] for r in con.execute("PRAGMA table_info(letters)").fetchall()}
    for c in ("kind", "replied_on", "photo"):
        if c not in cols:
            con.execute(f"ALTER TABLE letters ADD COLUMN {c} TEXT DEFAULT ''")
    con.commit()
    return con


def _now():
    return datetime.now().isoformat(timespec="seconds")


def archive_day(day, paper_csv=None, photo_glob=None):
    """Keep a copy of the paper list and the photo for this day, so you can open them later. Returns (paper_path, photo_path)."""
    folder = os.path.join(ARCHIVE, str(day))
    os.makedirs(folder, exist_ok=True)
    paper_out = photo_out = ""
    if paper_csv and os.path.exists(paper_csv):
        paper_out = os.path.join(folder, f"paper_{day}.csv")
        shutil.copy(paper_csv, paper_out)
    if photo_glob:
        photos = sorted(p for p in glob.glob(photo_glob) if p.lower().endswith((".jpg", ".jpeg", ".png")))
        if photos:
            for old in glob.glob(os.path.join(folder, "*")):          # replace the photos of an earlier check of this day
                if old.lower().endswith((".jpg", ".jpeg", ".png")):
                    os.remove(old)
            for p in photos:
                shutil.copy(p, os.path.join(folder, os.path.basename(p)))
            photo_out = os.path.join(folder, os.path.basename(photos[0]))
    return paper_out, photo_out


def save_check(day, rows, paper_csv=None, photo_glob=None, fingerprint_file="", attendance_file=""):
    """Save the result of a check (replaces an earlier check of the same day)."""
    paper_out, photo_out = archive_day(day, paper_csv, photo_glob)
    same = sum(1 for r in rows if not r["ask"])
    con = connect()
    with con:
        old = con.execute("SELECT filled_at, cells_written FROM days WHERE day=?", (str(day),)).fetchone()
        con.execute("DELETE FROM records WHERE day=?", (str(day),))
        con.execute("INSERT OR REPLACE INTO days VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (str(day), _now(), old["filled_at"] if old else None, old["cells_written"] if old else 0,
                     len(rows), same, len(rows) - same, paper_out, photo_out,
                     os.path.basename(fingerprint_file or ""), os.path.basename(attendance_file or "")))
        for r in rows:
            status = "same" if not r["ask"] else ("check" if r.get("ot_fill") is not None else "ask")
            fo = "" if r["Finger_OT"] == "" else f"{r['Finger_OT']:g}"
            po = "" if r["Paper_OT"] is None else f"{r['Paper_OT']:g}"
            ph = "" if r["Paper_Hours"] is None else f"{r['Paper_Hours']:g}"
            fill = "" if r.get("ot_fill") is None else f"{r['ot_fill']:g}"
            con.execute("INSERT INTO records VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (str(day), r["ID"], r["Name"], ph, po, r["In"], r["Out"], fo, r["Punches"], status,
                         "" if not r["ask"] else r["Action"], fill, 1 if r.get("night") else 0))
    con.close()


def mark_filled(day, n_cells):
    con = connect()
    with con:
        con.execute("UPDATE days SET filled_at=?, cells_written=cells_written+? WHERE day=?", (_now(), n_cells, str(day)))
    con.close()


def add_reason(day, worker_id, reason, final_ot=""):
    con = connect()
    with con:
        con.execute("INSERT INTO reasons(day, worker_id, reason, final_ot, saved_at) VALUES(?,?,?,?,?)",
                    (str(day), worker_id, reason, final_ot or "", _now()))
    con.close()


def save_decisions(day, dec):
    """What YOU decided for each worker of this day (hours, OT, code, reason). Saved again if you change it."""
    if not dec:
        return
    con = connect()
    with con:
        for pid, x in dec.items():
            con.execute("INSERT OR REPLACE INTO decisions(day, worker_id, hours, ot, code, reason, saved_at, le) VALUES(?,?,?,?,?,?,?,?)",
                        (str(day), pid, "" if x.get("hours") is None else f"{x['hours']:g}", "" if x.get("ot") is None else f"{x['ot']:g}",
                         x.get("code") or "", x.get("reason") or "", _now(), "" if not x.get("le") else str(int(x["le"]))))
    con.close()


def day_decisions(day):
    return _rows("SELECT * FROM decisions WHERE day=? ORDER BY worker_id", (str(day),))


def worker_decisions(pid):
    return _rows("SELECT * FROM decisions WHERE worker_id=? ORDER BY day DESC", (pid,))


def add_corrections(day, items):
    if not items:
        return
    con = connect()
    with con:
        for pid, field, was, now in items:
            con.execute("INSERT INTO corrections(day, worker_id, field, was, now, saved_at) VALUES(?,?,?,?,?,?)",
                        (str(day), pid, field, was, now, _now()))
    con.close()


def _rows(sql, args=()):
    con = connect()
    out = [dict(r) for r in con.execute(sql, args).fetchall()]
    con.close()
    return out


def days_list():
    return _rows("SELECT * FROM days ORDER BY day DESC")


def day_records(day):
    return _rows("SELECT * FROM records WHERE day=? ORDER BY worker_id", (str(day),))


def day_info(day):
    r = _rows("SELECT * FROM days WHERE day=?", (str(day),))
    return r[0] if r else None


def worker_ids():
    return [r["worker_id"] for r in _rows("SELECT DISTINCT worker_id FROM records ORDER BY worker_id")]


def worker_history(pid):
    recs = _rows("SELECT * FROM records WHERE worker_id=? ORDER BY day DESC", (pid,))
    reasons = _rows("SELECT day, reason, final_ot, saved_at FROM reasons WHERE worker_id=? ORDER BY day DESC, id DESC", (pid,))
    return recs, reasons


def month_remarks(year, month):
    prefix = f"{year:04d}-{month:02d}-"
    return _rows("SELECT day, worker_id, name, paper_ot, finger_ot, ot_filled, status, remark FROM records "
                 "WHERE day LIKE ? AND status != 'same' ORDER BY day, worker_id", (prefix + "%",))


def summary():
    d = _rows("SELECT COUNT(*) n FROM days")[0]["n"]
    r = _rows("SELECT COUNT(*) n FROM records")[0]["n"]
    w = _rows("SELECT COUNT(DISTINCT worker_id) n FROM records")[0]["n"]
    return d, r, w


def backup():
    os.makedirs("backup", exist_ok=True)
    dest = os.path.join("backup", f"attendance_db_{datetime.now():%Y%m%d_%H%M%S}.db")
    src = sqlite3.connect(DB)
    dst = sqlite3.connect(dest)
    with dst:
        src.backup(dst)
    src.close(); dst.close()
    return dest


# ---------------- missed punch letters (worker forgot to scan: they must write a letter) ----------------
def letters_for_day(day):
    return _rows("SELECT * FROM letters WHERE day=?", (str(day),))


def letters_sync(day, ticked, untick_ids):
    """ticked = list of dicts(worker_id, name, dept, problem, punches). Adds new ones (keeps what you already typed).
    untick_ids = workers you un-ticked: removed only if you have not yet told them (no dates typed)."""
    con = connect()
    with con:
        for t in ticked:
            con.execute("INSERT OR IGNORE INTO letters(day, worker_id, name, dept, problem, punches, kind, created_at) VALUES(?,?,?,?,?,?,?,?)",
                        (str(day), t["worker_id"], t.get("name", ""), t.get("dept", ""), t.get("problem", ""), t.get("punches", ""), t.get("kind", ""), _now()))
        for pid in untick_ids:
            con.execute("DELETE FROM letters WHERE day=? AND worker_id=? AND told_on='' AND returned_on='' AND replied_on=''", (str(day), pid))
    con.close()


def letters_list():
    return _rows("SELECT * FROM letters ORDER BY day DESC, worker_id")


def letters_update(items):
    """items = list of dicts(id, told_on, returned_on, note, delete)"""
    con = connect()
    with con:
        for x in items:
            if x.get("delete"):
                ph = con.execute("SELECT photo FROM letters WHERE id=?", (x["id"],)).fetchone()
                if ph and ph[0]:
                    _photo_cleanup([p for p in photo_list(ph[0]) if os.path.exists(p)])
                con.execute("DELETE FROM letters WHERE id=?", (x["id"],))
            else:
                con.execute("UPDATE letters SET told_on=?, returned_on=?, replied_on=?, kind=?, note=? WHERE id=?",
                            (x.get("told_on", ""), x.get("returned_on", ""), x.get("replied_on", ""), x.get("kind", ""), x.get("note", ""), x["id"]))
    con.close()


# ---------------- fingerprint machine problems (blackout, machine broken): so you remember the reason later ----------------
def outage_add(day, from_time, to_time, dept, note, to_day=""):
    note = " ".join(str(note or "").split())[:300]
    if not note:
        return "Type what happened (example: blackout, machine broken)."
    for t in (from_time, to_time):
        if t:
            try:
                datetime.strptime(t, "%H:%M")
            except ValueError:
                return "The time must look like 17:30."
    to_day = str(to_day or "")
    if to_day and to_day < str(day):
        return "The 'to' date is before the 'from' date."
    if to_day == str(day):
        to_day = ""
    con = connect()
    with con:
        con.execute("INSERT INTO fp_outages(day, from_time, to_time, dept, note, created_at, to_day) VALUES(?,?,?,?,?,?,?)",
                    (str(day), from_time or "", to_time or "", (dept or "").strip().upper(), note, _now(), to_day))
    con.close()
    return ""


def outages_for_day(day):
    return _rows("SELECT * FROM fp_outages WHERE day<=? AND (CASE WHEN to_day='' OR to_day IS NULL THEN day ELSE to_day END)>=? ORDER BY day, from_time, id", (str(day), str(day)))


def outages_list(limit=60):
    return _rows("SELECT * FROM fp_outages ORDER BY day DESC, from_time DESC, id DESC LIMIT ?", (int(limit),))


def outage_delete(oid):
    con = connect()
    with con:
        con.execute("DELETE FROM fp_outages WHERE id=?", (int(oid),))
    con.close()


# ---------------- the shift HR chose for a worker (Day / Night / Evening): remembered until HR changes it ----------------
SHIFTS = ("Day", "Night", "Evening")


def shift_prefs():
    return {r["worker_id"]: r["shift"] for r in _rows("SELECT worker_id, shift FROM worker_shift")}


def shift_set(worker_id, shift):
    """shift '' = forget (the app works it out from the punches again). Returns '' or an error text."""
    worker_id = str(worker_id or "").strip().upper()
    if not re.fullmatch(r"DMO-\d{5}", worker_id):
        return "Worker ID must look like DMO-10003."
    if shift not in SHIFTS and shift != "":
        return "Shift must be Day, Night or Evening."
    con = connect()
    with con:
        if shift == "":
            con.execute("DELETE FROM worker_shift WHERE worker_id=?", (worker_id,))
        else:
            con.execute("INSERT INTO worker_shift(worker_id, shift, updated) VALUES(?,?,?) ON CONFLICT(worker_id) DO UPDATE SET shift=excluded.shift, updated=excluded.updated",
                        (worker_id, shift, _now()))
    con.close()
    return ""


# ---------------- workers remembered as resigned (the paper said 'Resigned'): nothing is filled for them from that day on ----------------
def resigned_prefs():
    """{worker id: 'YYYY-MM-DD'}: the first day the worker was seen as resigned on the paper."""
    return {r["worker_id"]: r["since_day"] for r in _rows("SELECT * FROM resigned_workers")}


def resigned_set(worker_id, day):
    worker_id = str(worker_id or "").strip().upper()
    if not re.fullmatch(r"DMO-\d{5}", worker_id):
        return "Worker ID must look like DMO-10003."
    con = connect()
    with con:
        con.execute("INSERT INTO resigned_workers(worker_id, since_day, noted_at) VALUES(?,?,?) ON CONFLICT(worker_id) DO UPDATE SET "
                    "since_day=MIN(since_day, excluded.since_day)", (worker_id, str(day), _now()))
    con.close()
    return ""


def resigned_forget(worker_id):
    con = connect()
    with con:
        con.execute("DELETE FROM resigned_workers WHERE worker_id=?", (str(worker_id or "").strip().upper(),))
    con.close()


# ---------------- work times that are not normal (a worker who starts at 04:00 and ends at 13:00 ...) ----------------
def norm_time(v):
    """'4:00', '04:00', '4.00', '0400' -> '04:00'. '' -> ''. Not a time -> None."""
    v = str(v or "").strip()
    if not v:
        return ""
    m = re.fullmatch(r"(\d{1,2})[:.](\d{2})", v) or re.fullmatch(r"(\d{2})(\d{2})", v)
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    return f"{h:02d}:{mi:02d}" if h < 24 and mi < 60 else None


def times_prefs():
    """{worker id: ('04:00', '13:00')}: the work times remembered for a worker (every day until you change them)."""
    return {r["worker_id"]: (r["t_start"], r["t_end"]) for r in _rows("SELECT * FROM worker_times")}


def times_set(worker_id, t_start, t_end):
    """Remember the work times of a worker. Returns '' or an error text."""
    worker_id = str(worker_id or "").strip().upper()
    a, b = norm_time(t_start), norm_time(t_end)
    if not re.fullmatch(r"DMO-\d{5}", worker_id):
        return "Worker ID must look like DMO-10003."
    if not a or not b:
        return "Start and End must both be times like 04:00."
    ma, mb = int(a[:2]) * 60 + int(a[3:]), int(b[:2]) * 60 + int(b[3:])
    span = mb - ma if mb > ma else mb + 1440 - ma
    if a == b or span > 14 * 60:
        return "Start and End must be different, and the shift can be 14 hours long at most."
    con = connect()
    with con:
        con.execute("INSERT INTO worker_times(worker_id, t_start, t_end, updated) VALUES(?,?,?,?) ON CONFLICT(worker_id) DO UPDATE SET "
                    "t_start=excluded.t_start, t_end=excluded.t_end, updated=excluded.updated", (worker_id, a, b, _now()))
    con.close()
    return ""


def times_clear(worker_id):
    con = connect()
    with con:
        con.execute("DELETE FROM worker_times WHERE worker_id=?", (str(worker_id or "").strip().upper(),))
    con.close()


# ---------------- leave papers: late, leave early, day off ----------------
LEAVE_KINDS = ("Come late", "Leave early", "Day off", "Half day", "Other")
LEAVE_PHOTOS = os.path.join("memory", "leave_photos")


def photo_list(value):
    """The photo column keeps one or MORE file paths, joined with '|'. Returns the list."""
    return [p for p in str(value or "").split("|") if p.strip()]


def _photo_write(folder, prefix, name, data, n=0):
    """Writes one photo file. Returns (path, error text)."""
    ext = os.path.splitext(str(name or ""))[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".pdf"):
        return "", "The photo must be png, jpg or pdf."
    if len(data) > 8 * 1024 * 1024:
        return "", "The photo is bigger than 8 MB."
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{prefix}_{datetime.now():%H%M%S%f}_{n}{ext}")
    with open(path, "wb") as f:
        f.write(data)
    return path, ""


def _photo_cleanup(paths):
    for p in paths:
        try:
            os.remove(p)
        except OSError:
            pass


def leave_add(worker_id, day, kind, at_time, note, photo_name=None, photo_bytes=None, to_day="", code="", photos=None, medical=False):
    """photos = list of (file name, bytes): one paper can have several photos (leave letter + medical certificate ...)."""
    worker_id = str(worker_id or "").strip().upper()
    code = str(code or "").strip()
    err, to_day = _leave_check(worker_id, day, to_day, kind, at_time, code)
    if err:
        return err
    note = " ".join(str(note or "").split())[:300]
    files = list(photos or [])
    if photo_bytes:
        files.append((photo_name, photo_bytes))
    paths = []
    for n, (nm, data) in enumerate(files):
        p, e = _photo_write(LEAVE_PHOTOS, f"{worker_id}_{str(day)}", nm, data, n)
        if e:
            _photo_cleanup(paths)
            return e
        paths.append(p)
    con = connect()
    with con:
        con.execute("INSERT INTO leave_notes(worker_id, day, kind, at_time, note, photo, created_at, to_day, code, medical) VALUES(?,?,?,?,?,?,?,?,?,?)",
                    (worker_id, str(day), kind, at_time or "", note, "|".join(paths), _now(), to_day, code, "1" if medical else ""))
    con.close()
    return ""


def leave_add_photos(lid, photos):
    """Add more photos to a saved leave paper. photos = list of (name, bytes). Returns '' or an error text."""
    row = leave_get(lid)
    if not row:
        return "This leave paper is not in the list."
    paths = []
    for n, (nm, data) in enumerate(photos or []):
        p, e = _photo_write(LEAVE_PHOTOS, f"{row['worker_id']}_{row['day']}", nm, data, n)
        if e:
            _photo_cleanup(paths)
            return e
        paths.append(p)
    con = connect()
    with con:
        con.execute("UPDATE leave_notes SET photo=? WHERE id=?", ("|".join(photo_list(row.get("photo")) + paths), int(lid)))
    con.close()
    return ""


def leave_set_medical(lid, medical):
    con = connect()
    with con:
        con.execute("UPDATE leave_notes SET medical=? WHERE id=?", ("1" if medical else "", int(lid)))
    con.close()


def photo_remove(kind, rid, path):
    """Take one photo off a saved paper. kind 'L' = leave paper, 'M' = letter. The file itself stays in the folder, so Undo can bring it back."""
    table = "leave_notes" if kind == "L" else "letters"
    con = connect()
    try:
        row = con.execute(f"SELECT photo FROM {table} WHERE id=?", (int(rid),)).fetchone()
        have = photo_list(row[0] if row else "")
        if path not in have:
            return "This photo is not on this paper any more."
        with con:
            con.execute(f"UPDATE {table} SET photo=? WHERE id=?", ("|".join(p for p in have if p != path), int(rid)))
    finally:
        con.close()
    return ""


def _leave_check(worker_id, day, to_day, kind, at_time, code):
    """The same checks for a new and an edited leave. Returns (error, clean_to_day)."""
    if not re.fullmatch(r"DMO-\d{5}", worker_id):
        return "Worker ID must look like DMO-10003.", ""
    if kind not in LEAVE_KINDS:
        return "Choose what kind of paper it is.", ""
    if at_time:
        try:
            datetime.strptime(at_time, "%H:%M")
        except ValueError:
            return "The time must look like 15:00.", ""
    to_day = str(to_day or "")
    if to_day and to_day < str(day):
        return "The 'to' date is before the 'from' date.", ""
    if to_day == str(day):
        to_day = ""
    if kind == "Half day" and to_day:
        return "A half day is one day: 'To' must be the same day as 'From'.", ""
    if to_day and (datetime.strptime(to_day, "%Y-%m-%d") - datetime.strptime(str(day), "%Y-%m-%d")).days > 120:
        return "A leave of more than 120 days: please check the dates.", ""
    if code and code not in ("A", "HP", "HH", "PC", "M", "L1", "L2", "AL1", "AL2", "MC", "HW", "SW", "NP", "Holiday"):
        return "This code is not in the code list.", ""
    return "", to_day


def leave_get(lid):
    r = _rows("SELECT * FROM leave_notes WHERE id=?", (int(lid),))
    return r[0] if r else None


def leave_update(lid, worker_id, day, to_day, kind, at_time, note, code, medical=None):
    """Change a saved leave (the photo stays)."""
    worker_id = str(worker_id or "").strip().upper()
    code = str(code or "").strip()
    err, to_day = _leave_check(worker_id, day, to_day, kind, at_time, code)
    if err:
        return err
    note = " ".join(str(note or "").split())[:300]
    con = connect()
    with con:
        con.execute("UPDATE leave_notes SET worker_id=?, day=?, to_day=?, kind=?, at_time=?, note=?, code=? WHERE id=?",
                    (worker_id, str(day), to_day, kind, at_time or "", note, code, int(lid)))
        if medical is not None:
            con.execute("UPDATE leave_notes SET medical=? WHERE id=?", ("1" if medical else "", int(lid)))
    con.close()
    return ""


def leave_overlaps(worker_id, day, to_day, exclude_id=None):
    """Saved leaves of this worker that touch any day from `day` to `to_day`."""
    to_day = to_day or day
    rows = _rows("SELECT * FROM leave_notes WHERE worker_id=? AND day<=? AND (CASE WHEN to_day='' OR to_day IS NULL THEN day ELSE to_day END)>=? ORDER BY day",
                 (str(worker_id).upper(), str(to_day), str(day)))
    return [r for r in rows if exclude_id is None or r["id"] != exclude_id]


def leave_month(year, month, worker_id=None):
    """Leaves that touch this month (all workers, or one)."""
    first = f"{year:04d}-{month:02d}-01"
    last = f"{year:04d}-{month:02d}-31"
    q = "SELECT * FROM leave_notes WHERE day<=? AND (CASE WHEN to_day='' OR to_day IS NULL THEN day ELSE to_day END)>=?"
    a = [last, first]
    if worker_id:
        q += " AND worker_id=?"
        a.append(str(worker_id).upper())
    return _rows(q + " ORDER BY day, worker_id, id", tuple(a))


def leave_for_day(day):
    return _rows("SELECT * FROM leave_notes WHERE day<=? AND (CASE WHEN to_day='' OR to_day IS NULL THEN day ELSE to_day END)>=? ORDER BY worker_id, id", (str(day), str(day)))


def leave_recent(limit=40):
    return _rows("SELECT * FROM leave_notes ORDER BY day DESC, id DESC LIMIT ?", (int(limit),))


def leave_delete(lid):
    con = connect()
    with con:
        row = con.execute("SELECT photo FROM leave_notes WHERE id=?", (int(lid),)).fetchone()
        con.execute("DELETE FROM leave_notes WHERE id=?", (int(lid),))
    con.close()
    if row and row[0]:
        _photo_cleanup([p for p in photo_list(row[0]) if os.path.exists(p)])


# ---------------- photo of the missed-punch letter the worker wrote
LETTER_PHOTOS = os.path.join("memory", "letter_photos")


def letter_photo_save(letter_id, name, data):
    """Adds a photo to the letter (the ones saved before stay). Returns '' or an error text."""
    con = connect()
    try:
        row = con.execute("SELECT worker_id, day, photo FROM letters WHERE id=?", (int(letter_id),)).fetchone()
        if not row:
            return "This letter is not in the list."
        path, err = _photo_write(LETTER_PHOTOS, f"{row[0]}_{row[1]}", name, data)
        if err:
            return err
        with con:
            con.execute("UPDATE letters SET photo=? WHERE id=?", ("|".join(photo_list(row[2]) + [path]), int(letter_id)))
    finally:
        con.close()
    return ""


def letter_create(day, worker_id, name, dept, kind, returned_on, note=""):
    """Makes a missed-punch letter line from a photo (the worker already wrote the letter). Returns (letter id, error text).
    If a letter of this worker and day exists, it is kept and only the 'letter back' date is filled when empty."""
    worker_id = str(worker_id or "").strip().upper()
    if not re.fullmatch(r"DMO-\d{5}", worker_id):
        return None, "Worker ID must look like DMO-10003."
    try:
        datetime.strptime(str(day), "%Y-%m-%d")
    except ValueError:
        return None, "The missed-punch date is not valid."
    con = connect()
    try:
        with con:
            con.execute("INSERT OR IGNORE INTO letters(day, worker_id, name, dept, problem, punches, kind, created_at) VALUES(?,?,?,?,?,?,?,?)",
                        (str(day), worker_id, name or "", dept or "", "Letter photo added by you", "", kind or "", _now()))
            row = con.execute("SELECT id, returned_on FROM letters WHERE day=? AND worker_id=?", (str(day), worker_id)).fetchone()
            if row and not row[1]:
                con.execute("UPDATE letters SET returned_on=? WHERE id=?", (str(returned_on)[:10], row[0]))
            note = " ".join(str(note or "").split())[:300]
            if row and note:
                con.execute("UPDATE letters SET note=? WHERE id=? AND (note IS NULL OR note='')", (note, row[0]))
    finally:
        con.close()
    return (row[0] if row else None), ""


# ---------------------------------------------------------------- Step 3 draft (saved while you type, before you press the confirm button)
def draft_save(day, items):
    """items = {worker_id: {hours, ot, code, reason, le, letter}} with the text as typed. Replaces the draft of those workers."""
    if not items:
        return
    con = connect()
    try:
        con.execute("ALTER TABLE step3_draft ADD COLUMN shift TEXT DEFAULT ''")      # the shift you chose in the table (not remembered yet)
    except Exception:
        pass
    with con:
        for pid, x in items.items():
            con.execute("INSERT OR REPLACE INTO step3_draft(day, worker_id, hours, ot, code, reason, le, letter, saved_at, shift) VALUES(?,?,?,?,?,?,?,?,?,?)",
                        (str(day), pid, str(x.get("hours", "")), str(x.get("ot", "")), str(x.get("code", "")), str(x.get("reason", "")),
                         str(x.get("le", "")), 1 if x.get("letter") else 0, _now(), str(x.get("shift", ""))))
    con.close()


def draft_get(day):
    return {r["worker_id"]: r for r in _rows("SELECT * FROM step3_draft WHERE day=?", (str(day),))}


def draft_clear(day, ids):
    con = connect()
    with con:
        for pid in ids:
            con.execute("DELETE FROM step3_draft WHERE day=? AND worker_id=?", (str(day), pid))
    con.close()


# ---------------------------------------------------------------- Questions to ask (OT different, late/early, 1-2 punches, no fingerprint, not on paper)
CASE_KINDS = ["OT different", "Late / early", "1-2 punches", "No fingerprint", "Not on paper"]


def case_sync(day, items, kinds=None):
    """items = list of dicts(worker_id, kind, name, dept, shift, paper_ot, finger_ot, detail). New cases are added. A case you already asked keeps your answer.
    Details are refreshed while you have not asked. A case that is gone now (you fixed the paper) and was never asked is removed. Only the kinds in `kinds` are cleaned."""
    kinds = list(kinds or CASE_KINDS)
    con = connect()
    with con:
        for t in items:
            con.execute("INSERT OR IGNORE INTO worker_cases(day, worker_id, kind, name, dept, shift, paper_ot, finger_ot, detail, created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                        (str(day), t["worker_id"], t["kind"], t.get("name", ""), t.get("dept", ""), t.get("shift", ""), str(t.get("paper_ot", "")), str(t.get("finger_ot", "")),
                         t.get("detail", ""), _now()))
            con.execute("UPDATE worker_cases SET paper_ot=?, finger_ot=?, shift=?, detail=? WHERE day=? AND worker_id=? AND kind=? AND asked_on='' AND answer=''",
                        (str(t.get("paper_ot", "")), str(t.get("finger_ot", "")), t.get("shift", ""), t.get("detail", ""), str(day), t["worker_id"], t["kind"]))
        now = {(t["worker_id"], t["kind"]) for t in items}
        for r in con.execute("SELECT worker_id, kind FROM worker_cases WHERE day=? AND asked_on='' AND answer='' AND note=''", (str(day),)).fetchall():
            if r[1] in kinds and (r[0], r[1]) not in now:
                con.execute("DELETE FROM worker_cases WHERE day=? AND worker_id=? AND kind=?", (str(day), r[0], r[1]))
    con.close()


def case_list():
    return _rows("SELECT * FROM worker_cases ORDER BY day DESC, worker_id, kind")


def case_for_day(day):
    return _rows("SELECT * FROM worker_cases WHERE day=?", (str(day),))


def case_overdue(today, days=2):
    """Cases nobody asked about yet, found `days` or more days ago."""
    lim = (today - __import__("datetime").timedelta(days=days)).isoformat()
    return _rows("SELECT * FROM worker_cases WHERE asked_on='' AND answer='' AND substr(created_at,1,10) <= ? ORDER BY day, worker_id", (lim,))


def case_update(items):
    """items = list of dicts(id, asked_on, answer, note, delete)"""
    con = connect()
    with con:
        for x in items:
            if x.get("delete"):
                con.execute("DELETE FROM worker_cases WHERE id=?", (x["id"],))
            else:
                con.execute("UPDATE worker_cases SET asked_on=?, answer=?, note=? WHERE id=?", (x.get("asked_on", ""), x.get("answer", ""), x.get("note", ""), x["id"]))
    con.close()


def last_filled_day():
    """The day you filled in Excel most recently (the last time you pressed Step 4)."""
    r = _rows("SELECT day FROM days WHERE filled_at IS NOT NULL AND filled_at != '' ORDER BY filled_at DESC LIMIT 1")
    return r[0]["day"] if r else None


# ---------------- quick notes and reminders (what a leader told you, things to do later) ----------------
def note_add(text, worker_id="", due=""):
    text = str(text or "").strip()
    if not text:
        return "Write the note first."
    con = connect()
    with con:
        con.execute("INSERT INTO notes(created_at, worker_id, text, due) VALUES(?,?,?,?)", (_now(), str(worker_id or "").strip().upper(), text, str(due or "")))
    con.close()
    return ""


def notes_due(today):
    """Open notes with a date that has come (and are not snoozed past today)."""
    t = str(today)
    return _rows("SELECT * FROM notes WHERE done_at='' AND due<>'' AND due<=? AND (snooze_until='' OR snooze_until<=?) ORDER BY due, id", (t, t))


def notes_list(state="open", worker="", q=""):
    sql, a = "SELECT * FROM notes WHERE 1=1", []
    if state == "open":
        sql += " AND done_at=''"
    elif state == "done":
        sql += " AND done_at<>''"
    if worker:
        sql += " AND worker_id=?"
        a.append(str(worker).strip().upper())
    if q:
        sql += " AND (lower(text) LIKE ? OR lower(worker_id) LIKE ?)"
        a += [f"%{q.lower()}%"] * 2
    return _rows(sql + " ORDER BY CASE WHEN done_at='' THEN 0 ELSE 1 END, CASE WHEN due='' THEN '9999' ELSE due END, id DESC", tuple(a))


def notes_for_worker(worker_id):
    return _rows("SELECT * FROM notes WHERE worker_id=? ORDER BY id DESC", (str(worker_id).strip().upper(),))


def note_done(nid, done=True):
    con = connect()
    with con:
        con.execute("UPDATE notes SET done_at=? WHERE id=?", (_now() if done else "", int(nid)))
    con.close()


def note_snooze(nid, until):
    con = connect()
    with con:
        con.execute("UPDATE notes SET snooze_until=? WHERE id=?", (str(until), int(nid)))
    con.close()


def note_delete(nid):
    con = connect()
    with con:
        con.execute("DELETE FROM notes WHERE id=?", (int(nid),))
    con.close()


# ---------------- notes printed under a worker's row on the Worker list ----------------
def list_note_add(worker_id, day, text):
    text = str(text or "").strip()
    if not text:
        return "Write the note first."
    con = connect()
    with con:
        con.execute("INSERT INTO list_notes(worker_id, day, text, created_at) VALUES(?,?,?,?)", (str(worker_id).strip(), str(day), text, _now()))
    con.close()
    return ""


def list_notes_month(year, month):
    return _rows("SELECT * FROM list_notes WHERE substr(day,1,7)=? ORDER BY day, id", (f"{year}-{month:02d}",))


def list_note_delete(nid):
    con = connect()
    with con:
        con.execute("DELETE FROM list_notes WHERE id=?", (int(nid),))
    con.close()


# ---------------- notebook pages (like Notion: title + free text, folder, tags, worker, pinned, photos) ----------------
PAGE_DIR = os.path.join("memory", "page_files")


def page_new(title="Untitled", folder=""):
    con = connect()
    with con:
        cur = con.execute("INSERT INTO pages(title, folder, created_at, updated_at) VALUES(?,?,?,?)", (title, folder, _now(), _now()))
        pid = cur.lastrowid
    con.close()
    return pid


def page_get(pid):
    r = _rows("SELECT * FROM pages WHERE id=?", (int(pid),))
    return r[0] if r else None


def page_list(q="", folder="", worker=""):
    sql, a = "SELECT id, title, folder, tags, worker_id, pinned, updated_at, substr(body,1,80) AS snippet FROM pages WHERE 1=1", []
    if folder:
        sql += " AND folder=?"
        a.append(folder)
    if worker:
        sql += " AND worker_id=?"
        a.append(worker.strip().upper())
    if q:
        sql += " AND (lower(title) LIKE ? OR lower(body) LIKE ? OR lower(tags) LIKE ? OR lower(worker_id) LIKE ? OR lower(folder) LIKE ?)"
        a += [f"%{q.lower()}%"] * 5
    return _rows(sql + " ORDER BY pinned DESC, updated_at DESC", tuple(a))


def page_folders():
    return [r["folder"] for r in _rows("SELECT DISTINCT folder FROM pages WHERE folder<>'' ORDER BY folder")]


def page_save(pid, **f):
    ok = {"title", "body", "folder", "tags", "worker_id", "pinned"}
    f = {k: v for k, v in f.items() if k in ok}
    if "worker_id" in f:
        f["worker_id"] = str(f["worker_id"]).strip().upper()
    if not f:
        return
    con = connect()
    with con:
        con.execute("UPDATE pages SET " + ", ".join(f"{k}=?" for k in f) + ", updated_at=? WHERE id=?", (*f.values(), _now(), int(pid)))
    con.close()


def page_delete(pid):
    for fl in page_files(pid):
        try:
            os.remove(fl["path"])
        except OSError:
            pass
    con = connect()
    with con:
        con.execute("DELETE FROM page_files WHERE page_id=?", (int(pid),))
        con.execute("DELETE FROM pages WHERE id=?", (int(pid),))
    con.close()


def page_files(pid):
    return _rows("SELECT * FROM page_files WHERE page_id=? ORDER BY id", (int(pid),))


def page_file_add(pid, name, data):
    os.makedirs(PAGE_DIR, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9._-]", "_", name)[-60:] or "file"
    path = os.path.join(PAGE_DIR, f"{int(pid)}_{datetime.now():%Y%m%d%H%M%S%f}_{safe}")
    with open(path, "wb") as f:
        f.write(data)
    con = connect()
    with con:
        con.execute("INSERT INTO page_files(page_id, name, path, created_at) VALUES(?,?,?,?)", (int(pid), name, path, _now()))
        con.execute("UPDATE pages SET updated_at=? WHERE id=?", (_now(), int(pid)))
    con.close()


def page_file_delete(fid):
    r = _rows("SELECT path FROM page_files WHERE id=?", (int(fid),))
    if r:
        try:
            os.remove(r[0]["path"])
        except OSError:
            pass
    con = connect()
    with con:
        con.execute("DELETE FROM page_files WHERE id=?", (int(fid),))
    con.close()


def pages_for_worker(wid):
    return _rows("SELECT id, title, folder, updated_at, substr(body,1,120) AS snippet FROM pages WHERE worker_id=? ORDER BY updated_at DESC", (str(wid).strip().upper(),))


def set_record_night(day, worker_id, night):
    """History: you say this worker was on the night (or day) shift that day."""
    con = connect()
    with con:
        con.execute("UPDATE records SET night=? WHERE day=? AND worker_id=?", (1 if night else 0, str(day), worker_id))
    con.close()
