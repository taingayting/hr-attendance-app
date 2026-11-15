"""Sample data for the public demo: 14 invented workers on Monday 5 October 2026.

No real worker is in here. Names are "Sample Worker 01" and so on, and IDs start with DMO-.
"""
import csv
import os
from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd

DAY = date(2026, 10, 5)
IDS = [f"DMO-{10001 + i}" for i in range(14)]
NAMES = [f"Sample Worker {i + 1:02d}" for i in range(14)]
# paper OT tick, last scan of the day, number of scans
PLAN = [("2", "18:02", 4), ("3", "18:03", 4), ("?", "19:31", 4), ("1", None, 0), ("2", "18:05", 3),
        ("3.5", "19:33", 4), ("", "16:02", 4), ("1", "17:04", 4), ("2", "18:00", 4), ("", "16:04", 4),
        ("4", "20:02", 4), ("2.5", "18:31", 4), ("1", "17:02", 4), ("3", "19:03", 4)]
HEADERS = ["L/E", "day hours", "night hours", "SW", "(a) day OT x1.5", "(a) night OT x2", "(b) holiday OT day",
           "(b) holiday OT night", "(c) OT 1.5", "(c) OT 2", "(d) Sunday OT day", "(N) Sunday OT night", "code"]


def seed(data_dir):
    work = os.path.join(data_dir, "web_work")
    for d in ("fingerprint", "attendance", "paper"):
        os.makedirs(os.path.join(work, d), exist_ok=True)

    # the monthly attendance Excel: sheet 人事考勤 with one 13-column block per day, and sheet CV
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "人事考勤"
    ws["A2"] = datetime(2026, 10, 1)
    ws["A7"], ws["B7"] = "ID", "Name"
    for r, (i, n) in enumerate(zip(IDS, NAMES), 8):
        ws.cell(r, 1, i)
        ws.cell(r, 2, n)
    for d in range(31):
        c = 14 + 13 * d
        ws.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + 12)
        ws.cell(5, c, datetime(2026, 10, d + 1))
        for k, h in enumerate(HEADERS):
            ws.cell(7, c + k, h)
    cv = wb.create_sheet("CV")
    for i, (w, n) in enumerate(zip(IDS, NAMES)):
        cv.append([i + 1, w, n, "", "", "", datetime(2025, 1, 6)] + [""] * 12)
    wb.save(os.path.join(work, "attendance", "October_attendance_sample.xlsx"))

    # the fingerprint export: 4 scans a day (in, lunch out, lunch in, out)
    rows = []
    for d in (DAY, DAY + timedelta(days=1)):
        for k, (w, n, (_, out, scans)) in enumerate(zip(IDS, NAMES, PLAN)):
            if d != DAY:
                out, scans = "16:03", 4
            if not out:
                continue
            times = [f"06:{50 + k % 9:02d}", "11:31", "12:24", out]
            if scans == 3:
                times = [times[0], times[1], out]
            for t in times:
                rows.append({"Department": "Sewing A", "Name": n, "No.": w,
                             "Date/Time": datetime.combine(d, datetime.strptime(t, "%H:%M").time()),
                             "ID Number": w, "VerifyCode": "FP"})
    pd.DataFrame(rows).to_excel(os.path.join(work, "fingerprint", "fingerprint_sample.xlsx"), index=False)

    # the paper list, as HR would type it from the daily sheet
    with open(os.path.join(work, "paper", f"paper_{DAY}.csv"), "w", newline="", encoding="utf-8-sig") as f:
        out = csv.writer(f)
        out.writerow(["ID", "HOURS", "OT"])
        for w, (ot, _, _) in zip(IDS, PLAN):
            out.writerow([w, "8", ot])
