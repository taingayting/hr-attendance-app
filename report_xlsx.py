"""
The daily Excel report. It is made after the check (nothing is saved into your attendance Excel by this).

Sheets:
  Summary         the day in numbers: how many workers, totals of work hours and overtime, late workers
  All workers     one row per worker: in/out, hours, overtime split exactly as it will go into the attendance Excel, remark
  Needs checking  only the workers you must check, with empty columns where you type the reason and the final OT

Every number here comes from the same check that fills your Excel, so the report and the Excel can never disagree.
"""
import os, re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

GREEN, YELLOW, RED, BLUE, HEAD = "E6F4EA", "FFF4CC", "FDE7E7", "E3F0FF", "3D3929"


def _text(v):
    """A text cell. Text that starts with = + - @ would be run as a formula by Excel, so it is stored as plain text."""
    return "" if v is None else str(v)


def _put(ws, row, col, value):
    c = ws.cell(row, col)
    if isinstance(value, str):
        c.value = value
        c.data_type = "s"                              # never a formula, even if the text starts with "="
    else:
        c.value = value
    return c


def _minutes(t):
    """'06:52' -> minutes since midnight; '03:40+1' -> next day."""
    m = re.match(r"^(\d{1,2}):(\d{2})(\+1)?$", str(t or ""))
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2)) + (1440 if m.group(3) else 0)


def status_of(r):
    if r.get("decided"):
        return "Checked by you"
    if not r["ask"]:
        return "Same"
    return "Check later" if r.get("ot_fill") is not None else "Ask"


def ot_used(r):
    """The OT that goes into the attendance Excel for this worker (same rule as the fill step)."""
    if not r.get("hours"):
        return None
    if r.get("ot_fill") is not None:
        return r["ot_fill"]
    return r["Paper_OT"] if r.get("ot_ok") else None


def split_ot(ot, night):
    """Same split as the fill step: day OT 0-2h x1.5, 2-6h '(c) OT 1.5', above 6h '(c) OT 2'; night OT all x2."""
    if not ot:
        return 0, 0, 0, 0
    if night:
        return 0, 0, 0, ot
    return min(ot, 2), min(max(ot - 2, 0), 4), max(ot - 6, 0), 0


def _row_values(r):
    ot = ot_used(r)
    a, c15, c2, n2 = split_ot(ot, r.get("night"))
    t_in, t_out = _minutes(r["In"]), _minutes(r["Out"])
    on_site = round((t_out - t_in) / 60, 1) if (t_in is not None and t_out is not None and t_out >= t_in) else ""
    npun = len(str(r["Punches"]).split()) if r["Punches"] else 0
    early = re.search(r"EARLY LEAVE (\d+) min", r["Action"] or "")
    fo = r["Finger_OT"]
    return {
        "Status": status_of(r), "ID": r["ID"], "Name": r["Name"], "Dept": r.get("Dept", ""), "Shift": "Night" if r.get("night") else ("Day" if r["In"] else ""),
        "Paper hours": "" if r["Paper_Hours"] is None else r["Paper_Hours"], "Paper OT": "?" if r["Paper_OT"] is None else r["Paper_OT"],
        "In": r["In"], "Out": r["Out"], "Time at factory (h)": on_site, "Punches": npun or "",
        "Late (min)": r.get("Late_min") or "", "Early (min)": int(early.group(1)) if early else "",
        "Hours to fill": r["hours"] if r.get("hours") else "", "Code": r.get("code") or "",
        "OT by fingerprint": "" if fo == "" else fo, "OT to fill": ot if ot else "",
        "(a) day OT x1.5": a or "", "(c) OT 1.5": c15 or "", "(c) OT 2": c2 or "", "(a) night OT x2": n2 or "",
        "Remark (what is different)": "" if not r["ask"] else r["Action"],
        "History": r["History"], "Reason (you type)": r.get("reason", ""), "Final OT if changed": "",
    }


ALL_COLS = ["Status", "ID", "Name", "Dept", "Shift", "Paper hours", "Paper OT", "In", "Out", "Time at factory (h)", "Punches",
            "Late (min)", "Early (min)", "Hours to fill", "Code", "OT by fingerprint", "OT to fill",
            "(a) day OT x1.5", "(c) OT 1.5", "(c) OT 2", "(a) night OT x2",
            "Remark (what is different)", "History", "Reason (you type)", "Final OT if changed"]
CHECK_COLS = ["Status", "ID", "Name", "Dept", "Paper hours", "Paper OT", "OT by fingerprint", "OT to fill", "In", "Out", "Late (min)", "Early (min)",
              "Remark (what is different)", "All punches", "History", "Reason (you type)", "Final OT if changed"]
WIDTH = {"Dept": 16, "No.": 6, "Status": 12, "ID": 12, "Name": 24, "Remark (what is different)": 60, "History": 36, "All punches": 40,
         "Reason (you type)": 28, "Final OT if changed": 14, "Time at factory (h)": 12}

_thin = Side(style="thin", color="D9D6CC")
_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _sheet(ws, cols, records):
    cols = ["No."] + list(cols)                        # row numbers, so you can say "row 7"
    records = [dict(r, **{"No.": i}) for i, r in enumerate(records, 1)]
    for j, c in enumerate(cols, 1):
        h = _put(ws, 1, j, c)
        h.font = Font(bold=True, color="FFFFFF")
        h.fill = PatternFill("solid", fgColor=HEAD)
        h.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = WIDTH.get(c, 10)
    ws.row_dimensions[1].height = 32
    for i, rec in enumerate(records, 2):
        color = {"Same": GREEN, "Check later": YELLOW, "Ask": RED, "Checked by you": BLUE}[rec["Status"]]
        for j, c in enumerate(cols, 1):
            cell = _put(ws, i, j, rec.get(c, ""))
            cell.fill = PatternFill("solid", fgColor=color)
            cell.border = _border
            cell.alignment = Alignment(wrap_text=c in ("Remark (what is different)", "History", "All punches"), vertical="top")
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(len(records) + 1, 2)}"


def build_workbook(day, rows):
    recs = [_row_values(r) for r in rows]
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    n = len(rows)
    same = sum(1 for x in recs if x["Status"] == "Same")
    later = sum(1 for x in recs if x["Status"] == "Check later")
    ask = sum(1 for x in recs if x["Status"] == "Ask")
    mine = sum(1 for x in recs if x["Status"] == "Checked by you")
    filled = [r for r in rows if r.get("hours")]
    day_h = sum(float(r["hours"]) for r in filled if not r.get("night"))
    night_h = sum(float(r["hours"]) for r in filled if r.get("night"))
    tot = [0.0, 0.0, 0.0, 0.0]
    for r in filled:
        for k, v in enumerate(split_ot(ot_used(r), r.get("night"))):
            tot[k] += v
    late_n = sum(1 for r in rows if r.get("Late_min"))
    early_n = sum(1 for x in recs if x["Early (min)"] != "")
    lines = [
        ("Attendance report", f"{day:%A %d %B %Y}"),
        ("Made", f"{datetime.now():%Y-%m-%d %H:%M}"),
        ("", ""),
        ("Workers on the paper", n),
        ("Same (paper = fingerprint)", same),
        ("Check later (I filled the fingerprint OT)", later),
        ("Ask (not filled)", ask),
        ("Checked by you (your numbers are used)", mine),
        ("", ""),
        ("Workers with hours to fill", len(filled)),
        ("Day hours to fill (total)", day_h),
        ("Night hours to fill (total)", night_h),
        ("(a) day OT x1.5 (total hours)", tot[0]),
        ("(c) OT 1.5 (total hours)", tot[1]),
        ("(c) OT 2 (total hours)", tot[2]),
        ("(a) night OT x2 (total hours)", tot[3]),
        ("Total OT hours", sum(tot)),
        ("", ""),
        ("Late workers (more than 3 min)", late_n),
        ("Early leavers (more than 3 min)", early_n),
        ("", ""),
        ("How to read this", "The fingerprint is the reference. Green rows are the same as the paper. Yellow rows: the fingerprint OT was used and the "
                             "difference is remarked for you to check. Red rows are NOT filled: ask the worker or supervisor, then type the reason on the sheet 'Needs checking'."),
        ("Numbers", "Hours and OT here are exactly what the fill step writes into your attendance Excel (it never overwrites a cell that already has a value)."),
    ]
    for i, (a, b) in enumerate(lines, 1):
        ca = _put(ws, i, 1, a)
        cb = _put(ws, i, 2, b)
        ca.font = Font(bold=True)
        cb.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        if isinstance(b, float):
            cb.number_format = "0.0"
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 90

    w2 = wb.create_sheet("All workers")
    _sheet(w2, ALL_COLS, recs)

    w3 = wb.create_sheet("Needs checking")
    need = []
    for r, x in zip(rows, recs):
        if r["ask"]:
            y = dict(x)
            y["All punches"] = r["Punches"]
            need.append(y)
    order = {"Ask": 0, "Check later": 1, "Checked by you": 2}
    need.sort(key=lambda y: order.get(y["Status"], 2))
    _sheet(w3, CHECK_COLS, need)
    return wb


def write_report(day, rows, path):
    """Build the workbook and save it. Returns the path."""
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    build_workbook(day, rows).save(path)
    return path