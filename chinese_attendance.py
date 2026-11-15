"""Chinese workers: fill the attendance Excel (sheet 'Day shift') from the Chinese fingerprint export.

Rules (all decided by HR):
  * Day and night shift are NOT separated: the total goes in column '白班上班时数 / day hours'.
  * Hours of one work day = (last punch - first punch) - 1 hour (lunch or night break), rounded to the nearest 0.5.
  * Only ONE punch = 10.5.   No punch on a working day = code A.
  * Sunday: the total goes in column (d) OT 2 (周日加班). Public holiday: column (b) holiday OT.  No punch on Sunday / holiday = nothing.
  * A night shift belongs to the day it STARTED.
  * Not sure (very short or very long day, two work sessions in one day, name or ID does not match ...) = I ASK. Nothing is guessed.
  * A cell that already has something is never overwritten.
"""
import os, re, math, shutil, difflib
from copy import copy
from datetime import datetime, date, timedelta

import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator

SHEET, CV_SHEET = "Day shift", "CV"
LUNCH_H = 1.0              # lunch (day) or break (night) taken off
SINGLE_H = 10.5            # only one punch
GAP_H = 8.0                # a pause of 8 hours or more = a new work session
SHORT_H = 7.0              # a session shorter than this (but longer than 1 hour) is asked
LONG_H = 17.0              # a session longer than this is asked
KIND_COL = {"day": "dayhours", "sun": "(d)ot2", "hol": "(b)holidayot", "code": "code"}


def nid(s):
    """DMO-05 -> DMO05, DMO-001 -> DMO001, DMO064 -> DMO064. (Keeps the number of digits: DMO05 and DMO-005 are different people.)"""
    return re.sub(r"[^A-Z0-9]", "", str(s or "").upper())


def nname(s):
    return re.sub(r"[^A-Z]", "", str(s or "").upper())


def std_id(s):
    """The ID written for a NEW worker: DMO-056."""
    m = re.search(r"(\d+)\s*$", str(s or ""))
    return f"DMO-{int(m.group(1)):03d}" if m else str(s).strip()


def norm_head(v):
    return re.sub(r"\s+", "", str(v or "")).lower()


def round_half(x):
    return math.floor(x * 2 + 0.5) / 2


def _num(v):
    return int(v) if float(v) == int(v) else float(v)


def _to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for f in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(str(v).strip(), f).date()
        except ValueError:
            pass
    return None


# ============================================================ fingerprint
def read_fp(path):
    df = pd.read_excel(path)
    def col(*names):
        for n in names:
            for c in df.columns:
                if n.lower() in str(c).lower():
                    return c
        raise ValueError(f"The fingerprint file has no column like '{names[0]}'. I found: {', '.join(map(str, df.columns))}")
    out = pd.DataFrame({"id": df[col("ID Number", "ID")].astype(str).str.strip(), "name": df[col("Name")].astype(str).str.strip(),
                        "dt": pd.to_datetime(df[col("Date/Time", "Date")], errors="coerce")}).dropna(subset=["dt"])
    return out.sort_values(["id", "dt"]).reset_index(drop=True)


def build_sessions(ts):
    """Punches (sorted datetimes) -> work sessions. A pause of 8 h or more starts a new one. A lone morning punch followed by a lone evening punch
    of the same date (admin staff who punch only in and out) is ONE day session."""
    ts = sorted(ts)
    if not ts:
        return []
    out = [[ts[0]]]
    for t in ts[1:]:
        if (t - out[-1][-1]).total_seconds() / 3600 >= GAP_H:
            out.append([t])
        else:
            out[-1].append(t)
    res, i = [], 0
    while i < len(out):
        s = out[i]
        if i + 1 < len(out) and 4 <= s[0].hour <= 10 and (len(s) == 1 or len(s) == 2 and s[-1].hour <= 13):
            n = out[i + 1]
            if 15 <= n[0].hour <= 21 and n[0].date() == s[0].date() and (n[-1] - s[0]).total_seconds() / 3600 <= 14.0:
                res.append(s + n)
                i += 2
                continue
        res.append(s)
        i += 1
    return res


def session_hours(s):
    """-> (value or None, ask_text or '', note, suggestion)"""
    if len(s) == 1:
        return SINGLE_H, "", "1 punch = 10.5", None
    span = (s[-1] - s[0]).total_seconds() / 3600
    if span < 1:
        return SINGLE_H, "", "2+ punches within 1 hour = counted as 1 punch = 10.5", None
    sug = round_half(span - LUNCH_H) if span >= 2 else None
    if s[0].hour < 5:
        return None, "Starts after midnight with no earlier punch: maybe the end of last night's shift (the day before)", "", sug
    if span < SHORT_H:
        return None, f"Short day: {span:.1f} h from first to last punch", "", sug
    if span > LONG_H:
        return None, f"Very long: {span:.1f} h from first to last punch (two shifts in a row?)", "", sug
    return round_half(span - LUNCH_H), "", "", None


# ============================================================ the workbook
def read_book(path):
    """Everything we need from the Excel (formulas are NOT calculated; dates come from A3 and A2). Returns (wb, info)."""
    wb = openpyxl.load_workbook(path)
    if SHEET not in wb.sheetnames or CV_SHEET not in wb.sheetnames:
        raise ValueError(f"This Excel has no sheets '{SHEET}' and '{CV_SHEET}'. Is it the Chinese attendance file?")
    ws, cv = wb[SHEET], wb[CV_SHEET]
    start, end = _to_date(ws["A3"].value), _to_date(ws["A2"].value)
    if not start or not end:
        raise ValueError("I cannot read the month: cells A3 (first day) and A2 (last day) of sheet 'Day shift' must be dates.")
    blocks = sorted([m for m in ws.merged_cells.ranges if m.min_row == 5 and m.min_col >= 15], key=lambda m: m.min_col)
    days = {}
    for k, b in enumerate(blocks):
        d = start + timedelta(days=k)
        if d > end:
            break
        cols = {norm_head(ws.cell(7, c).value): c for c in range(b.min_col, b.max_col + 1)}
        for need in ("dayhours", "(d)ot2", "(b)holidayot", "code"):
            if need not in cols:
                raise ValueError(f"The block of {d:%d %b} has no column '{need}'. Did the headers change?")
        days[d] = cols
    cvw = {}
    for r in range(3, cv.max_row + 1):
        i = cv.cell(r, 2).value
        if i:
            cvw[str(i).strip()] = {"row": r, "zh": cv.cell(r, 5).value, "name": cv.cell(r, 6).value or "", "group": cv.cell(r, 3).value, "title": cv.cell(r, 4).value,
                                   "entry": _to_date(cv.cell(r, 7).value), "stop": _to_date(cv.cell(r, 19).value)}
    day_rows = {}
    for r in range(8, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if v:
            day_rows[str(v).strip()] = r
    return wb, {"start": start, "end": end, "days": days, "cv": cvw, "day_rows": day_rows}


# ============================================================ matching fingerprint people with the Excel
def match_people(fp, info):
    """One row per fingerprint person: status 'ok' | 'name_diff' | 'by_name' | 'new', and the suggested Excel ID."""
    cv = info["cv"]
    by_id = {nid(i): i for i in cv}
    by_nm = {}
    for i, v in cv.items():
        if nname(v["name"]):
            by_nm.setdefault(nname(v["name"]), []).append(i)
    rows = []
    for pid, g in fp.groupby("id"):
        name = g["name"].iloc[0]
        t = by_id.get(nid(pid))
        if t:
            sim = difflib.SequenceMatcher(None, nname(name), nname(cv[t]["name"])).ratio() if cv[t]["name"] else 1.0
            st = "ok" if sim >= 0.6 else "name_diff"
            rows.append({"fp_id": pid, "name": name, "status": st, "target": t, "excel_name": cv[t]["name"], "first": g["dt"].min(), "last": g["dt"].max(), "n": len(g)})
            continue
        t2 = by_nm.get(nname(name))
        if t2:
            rows.append({"fp_id": pid, "name": name, "status": "by_name", "target": t2[0], "excel_name": cv[t2[0]]["name"], "first": g["dt"].min(), "last": g["dt"].max(), "n": len(g)})
        else:
            rows.append({"fp_id": pid, "name": name, "status": "new", "target": std_id(pid), "excel_name": "", "first": g["dt"].min(), "last": g["dt"].max(), "n": len(g)})
    return rows


# ============================================================ the plan
def make_plan(fp, info, people, decisions, d_from, d_to, day_type, both_choice=None):
    """decisions = {fp_id: Excel ID | '' (skip)}. day_type(d) -> 'normal' | 'sunday' | 'holiday' | 'both'.
    Returns dict(cells=[...], asks=[...], notes=[...])."""
    cv, days = info["cv"], info["days"]
    cells, asks, skipped = [], [], []
    by_pid = {p["fp_id"]: p for p in people}
    fp_first, fp_last = fp["dt"].min(), fp["dt"].max()
    done_ids = set()
    for pid, g in fp.groupby("id"):
        target = decisions.get(pid, by_pid[pid]["target"] if pid in by_pid else "")
        if not target or target in done_ids:
            continue
        done_ids.add(target)
        p = by_pid.get(pid, {})
        c = cv.get(target, {})
        entry = c.get("entry") or (g["dt"].min().date() if target not in cv else None)
        stop = c.get("stop")
        sess = build_sessions(list(g["dt"]))
        by_day = {}
        for s in sess:
            by_day.setdefault(s[0].date(), []).append(s)
        for d in days:
            if d < d_from or d > d_to:
                continue
            if entry and d < entry:
                continue
            if stop and d >= stop:
                continue
            dtp = day_type(d)
            ss = by_day.get(d, [])
            base = {"cid": target, "name": p.get("name") or c.get("name", ""), "day": d, "dtype": dtp}
            if not ss:
                if dtp == "normal" and d >= fp_first.date() and d <= fp_last.date():
                    cells.append({**base, "kind": "code", "value": "A", "note": "no punch", "ask": ""})
                continue
            if len(ss) > 1:
                asks.append({**base, "kind": "day" if dtp == "normal" else "sun" if dtp == "sunday" else "hol", "ask": f"{len(ss)} work sessions start on this day", "suggest": None,
                             "punches": " | ".join(" ".join(x.strftime("%H:%M") for x in s) for s in ss)})
                continue
            s = ss[0]
            val, ask, note, sug = session_hours(s)
            kind = {"normal": "day", "sunday": "sun", "holiday": "hol"}.get(dtp)
            txt = " ".join(x.strftime("%a %H:%M") if x.date() != s[0].date() else x.strftime("%H:%M") for x in s)
            if dtp == "both":
                asks.append({**base, "kind": "both", "ask": "Sunday AND public holiday: which column?", "suggest": val if val is not None else sug, "punches": txt})
                continue
            if val is None:
                asks.append({**base, "kind": kind, "ask": ask, "suggest": sug, "punches": txt})
            else:
                cells.append({**base, "kind": kind, "value": val, "note": note, "ask": "", "punches": txt})
    return {"cells": cells, "asks": asks}


def grid(plan, info, d_from, d_to):
    """Workers x days table of what will be written (for the screen)."""
    days = [d for d in info["days"] if d_from <= d <= d_to]
    rows = {}
    for c in plan["cells"]:
        r = rows.setdefault(c["cid"], {"ID": c["cid"], "Name": c["name"]})
        r[f"{c['day']:%d}"] = ("A" if c["kind"] == "code" else (f"{_num(c['value'])}" + {"day": "", "sun": " (Sun)", "hol": " (Hol)"}[c["kind"]]))
    for a in plan["asks"]:
        r = rows.setdefault(a["cid"], {"ID": a["cid"], "Name": a["name"]})
        r[f"{a['day']:%d}"] = "❓"
    cols = ["ID", "Name"] + [f"{d:%d}" for d in days]
    df = pd.DataFrame(list(rows.values())).reindex(columns=cols).fillna("")
    return df.sort_values("ID").reset_index(drop=True)


# ============================================================ write
def _copy_row_style(ws, src, dst, c1, c2):
    for c in range(c1, c2 + 1):
        a, b = ws.cell(src, c), ws.cell(dst, c)
        if a.has_style:
            b._style = copy(a._style)
        if isinstance(a.value, str) and a.value.startswith("="):
            b.value = Translator(a.value, origin=a.coordinate).translate_formula(b.coordinate)


def apply_changes(path, writes, new_workers, out_path, backup_dir="backup"):
    """writes = [{cid, day, kind, value}]  new_workers = [{id, name, group, title, entry}].
    Saves to out_path (a backup of the original is made first), reads every cell back. Returns (log, n_written, bad)."""
    wb, info = read_book(path)
    ws, cv = wb[SHEET], wb[CV_SHEET]
    log = []
    # 1) new workers -> sheet CV (next free row, same look and formulas as the row above)
    last = max([v["row"] for v in info["cv"].values()] or [2])
    for w in new_workers:
        if w["id"] in info["cv"]:
            continue
        last += 1
        _copy_row_style(cv, last - 1, last, 1, cv.max_column)
        cv.cell(last, 1).value = cv.cell(last - 1, 1).value
        cv.cell(last, 2).value = w["id"]
        cv.cell(last, 3).value = w.get("group") or None
        cv.cell(last, 4).value = w.get("title") or None
        cv.cell(last, 6).value = w.get("name")
        if w.get("entry"):
            cv.cell(last, 7).value = datetime.combine(w["entry"], datetime.min.time())
        info["cv"][w["id"]] = {"row": last}
        log.append(f"NEW WORKER added to sheet CV row {last}: {w['id']} {w.get('name')}")
    # 2) everybody who gets a cell must be in the list of sheet 'Day shift'
    day_rows = dict(info["day_rows"])
    nxt = max(day_rows.values() or [7]) + 1
    for cid in sorted({x["cid"] for x in writes}):
        if cid not in day_rows:
            ws.cell(nxt, 2).value = cid
            day_rows[cid] = nxt
            log.append(f"ADDED to the list of sheet 'Day shift' row {nxt}: {cid}")
            nxt += 1
    # 3) the cells
    changes, skipped = [], 0
    for x in writes:
        cols = info["days"].get(x["day"])
        row = day_rows.get(x["cid"])
        if not cols or not row:
            log.append(f"SKIPPED {x['cid']} {x['day']}: no column or no row")
            continue
        cell = ws.cell(row, cols[KIND_COL[x["kind"]]])
        v = x["value"] if x["kind"] == "code" else _num(x["value"])
        if cell.value == v:
            continue
        if cell.value not in (None, ""):
            skipped += 1
            log.append(f"SKIPPED {cell.coordinate} ({x['cid']} {x['day']:%d %b}): already has {cell.value!r}, wanted {v!r}")
            continue
        cell.value = v
        changes.append((cell.coordinate, v))
    os.makedirs(backup_dir, exist_ok=True)
    shutil.copy(path, os.path.join(backup_dir, os.path.basename(path).replace(".xlsx", f"_{datetime.now():%Y%m%d_%H%M%S}.xlsx")))
    wb.save(out_path)
    chk = openpyxl.load_workbook(out_path)[SHEET]
    bad = [(c, v, chk[c].value) for c, v in changes if chk[c].value != v]
    return log, len(changes), bad
