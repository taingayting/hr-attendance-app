"""Page 'Monthly report': the report for the boss. Who was late / early, who was absent, who missed punches, where the paper and the fingerprint disagree,
leave, OT, per department, compared with last month. Dashboard on screen + Excel file (one sheet per topic) + one printable page.
Numbers come from the attendance Excel of the month and from the app memory (letters, questions). Nothing here changes the Excel."""
import html as _html
import io
import os
from datetime import date

import pandas as pd

import scorecard as SC

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
# for these numbers MORE is WORSE (arrows are red when they go up)
WORSE_UP = {"Absent days (A)", "Late / early incidents", "Late / early minutes", "Missed scans", "Paper ≠ fingerprint", "Bad workers", "Watch workers"}


def _present_days(loaded):
    n = 0
    for per in loaded["data"].values():
        for x in per.values():
            if (x.get("day") or 0) > 0 or (x.get("night") or 0) > 0 or (x.get("hol") or 0) > 0 or (x.get("sun") or 0) > 0:
                n += 1
    return n


def build(loaded, year, month, letters, cases, outages, rules):
    df, det = SC.compute(loaded, year, month, letters, cases, outages, rules)
    name = {r.ID: (r.English or r.Name) for r in df.itertuples()}
    dept = {r.ID: r.Department for r in df.itertuples()}
    ym = f"{year}-{month:02d}"
    # ---- one table per topic (a line = one worker and one day)
    topics = {"late": [], "absent": [], "missed": [], "paper": [], "nofp": [], "notpaper": []}
    trend = {}
    for wid, items in det.items():
        for d, what in items:
            key = None
            if what.startswith("Late/early"):
                key = "late"
            elif what.startswith("Absent"):
                key = "absent"
            elif what.startswith("Missed scan"):
                key = "missed"
            elif what.startswith("OT different"):
                key = "paper"
            elif what.startswith("No fingerprint"):
                key = "nofp"
            elif what.startswith("Not on paper"):
                key = "notpaper"
            if key:
                topics[key].append({"Worker": wid, "Name": name.get(wid, ""), "Department": dept.get(wid, ""), "Day": d, "What": what})
                t = trend.setdefault(d, {"late": 0, "absent": 0, "missed": 0, "paper": 0})
                if key in t:
                    t[key] += 1
    tdf = {k: pd.DataFrame(v, columns=["Worker", "Name", "Department", "Day", "What"]).sort_values(["Day", "Worker"]) if v else
           pd.DataFrame(columns=["Worker", "Name", "Department", "Day", "What"]) for k, v in topics.items()}
    # ---- OT hours and leave
    ot_by_dept, leave_rows = {}, []
    leave_days = 0.0
    for w in loaded["workers"]:
        per = loaded["data"].get(w["id"], {})
        tot = sum((x.get("ot") or 0) + (x.get("extra") or 0) + (x.get("hol") or 0) + (x.get("sun") or 0) for x in per.values())
        ot_by_dept[w["dept"]] = ot_by_dept.get(w["dept"], 0) + tot
    leave_days = float(df["Leave days"].sum())
    present = _present_days(loaded)
    absent = int(df["Absent (A)"].sum())
    k = {
        "Workers": len(df),
        "Attendance rate %": round(100 * present / (present + absent), 1) if (present + absent) else 0.0,
        "Absent days (A)": absent,
        "Late / early incidents": int(df["Late days"].sum()),
        "Late / early minutes": float(df["Late/early min"].sum()),
        "Missed scans": int(df["Missed scans"].sum()),
        "Paper ≠ fingerprint": int(df["OT questions"].sum()),
        "Leave days": leave_days,
        "OT hours (all kinds)": round(sum(ot_by_dept.values()), 1),
        "Bad workers": int((df["Result"] == "Bad").sum()),
        "Watch workers": int((df["Result"] == "Watch").sum()),
        "Good workers": int((df["Result"] == "Good").sum()),
    }
    dep = SC.by_department(df)
    dep["OT hours"] = dep["Department"].map(lambda d: round(ot_by_dept.get(d, 0), 1))
    tr = pd.DataFrame([{"Day": d, "Late / early": v["late"], "Absent": v["absent"], "Missed scans": v["missed"], "Paper ≠ fingerprint": v["paper"]} for d, v in sorted(trend.items())])
    return {"ym": (year, month), "title": f"{MONTHS[month - 1]} {year}", "df": df, "det": det, "kpi": k, "dept": dep, "topics": tdf, "trend": tr, "ot_by_dept": ot_by_dept}


def arrow(name, now, before):
    if before is None:
        return "", ""
    d = now - before
    if abs(d) < 1e-9:
        return "＝", "#666"
    up = d > 0
    worse = (up and name in WORSE_UP) or (not up and name not in WORSE_UP and name in ("Attendance rate %", "Good workers"))
    better = (not up and name in WORSE_UP) or (up and name in ("Attendance rate %", "Good workers"))
    col = "#b91c1c" if worse else "#15803d" if better else "#666"
    return f"{'▲' if up else '▼'} {abs(d):g}", col


def repeat_offenders(reps):
    """Workers flagged (Bad or Watch) in more than one of the given months."""
    cnt, info = {}, {}
    for r in reps:
        for x in r["df"].itertuples():
            if x.Result in ("Bad", "Watch"):
                cnt.setdefault(x.ID, []).append(f"{MONTHS[r['ym'][1] - 1][:3]} {x.Result}")
                info[x.ID] = (x.English or x.Name, x.Department)
    rows = [{"Worker": w, "Name": info[w][0], "Department": info[w][1], "Months flagged": len(v), "History": ", ".join(v)} for w, v in cnt.items() if len(v) >= 2]
    return pd.DataFrame(rows).sort_values(["Months flagged", "Worker"], ascending=[False, True]) if rows else pd.DataFrame(columns=["Worker", "Name", "Department", "Months flagged", "History"])


def worst(rep, n=10):
    return rep["df"].sort_values(["_score", "ID"], ascending=[False, True]).head(n).drop(columns=["_score"])


# ------------------------------------------------------------------ files
def to_xlsx(rep, prev, offenders):
    buf = io.BytesIO()
    k = rep["kpi"]
    rows = []
    for n_, v in k.items():
        b = prev["kpi"].get(n_) if prev else None
        a_, _c = arrow(n_, v, b)
        rows.append({"Number": n_, rep["title"]: v, (prev["title"] if prev else "Last month"): "" if b is None else b, "Change": a_})
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(rows).to_excel(xw, sheet_name="Summary", index=False)
        rep["dept"].to_excel(xw, sheet_name="By department", index=False)
        worst(rep, 15).to_excel(xw, sheet_name="Worst 15 workers", index=False)
        rep["df"].sort_values("_score", ascending=False).drop(columns=["_score"]).to_excel(xw, sheet_name="All workers", index=False)
        for key, sheet in (("late", "Late or early"), ("absent", "Absent (A)"), ("missed", "Missed punch"), ("paper", "Paper vs fingerprint OT"),
                           ("nofp", "On paper no fingerprint"), ("notpaper", "Fingerprint not on paper")):
            rep["topics"][key].to_excel(xw, sheet_name=sheet, index=False)
        offenders.to_excel(xw, sheet_name="Repeat offenders", index=False)
        if len(rep["trend"]):
            rep["trend"].to_excel(xw, sheet_name="Day by day", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                w = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(60, max(10, w + 2))
            for c in ws[1]:
                c.font = c.font.copy(bold=True)
    return buf.getvalue()


def to_html(rep, prev, offenders):
    esc = _html.escape
    k = rep["kpi"]
    cards = ""
    for n_ in ("Attendance rate %", "Absent days (A)", "Late / early incidents", "Late / early minutes", "Missed scans", "Paper ≠ fingerprint", "Leave days", "OT hours (all kinds)"):
        a_, col = arrow(n_, k[n_], prev["kpi"].get(n_) if prev else None)
        cards += f"<div class='c'><div class='t'>{esc(n_)}</div><div class='v'>{k[n_]:g}</div><div class='a' style='color:{col}'>{esc(a_)}</div></div>"

    def table(df, cols=None, limit=None):
        d = df if cols is None else df[cols]
        if limit:
            d = d.head(limit)
        h = "<tr>" + "".join(f"<th>{esc(str(c))}</th>" for c in d.columns) + "</tr>"
        body = ""
        for r in d.itertuples(index=False):
            body += "<tr>" + "".join(f"<td>{esc(str(v))}</td>" for v in r) + "</tr>"
        return f"<table>{h}{body}</table>"
    w = worst(rep, 10)[["ID", "English", "Department", "Result", "Why"]]
    dep = rep["dept"][["Department", "Workers", "Bad", "Watch", "Good", "% Good", "Missed scans", "Late/early min", "Absent (A)", "OT hours"]]
    off = table(offenders, limit=10) if len(offenders) else "<p>None.</p>"
    return f"""<!doctype html><html><head><meta charset="utf-8"><title>Attendance report {esc(rep['title'])}</title>
<style>@page{{size:A4;margin:12mm}}body{{font-family:Segoe UI,Arial,sans-serif;color:#222;margin:0;padding:14px;font-size:12px}}
h1{{font-size:20px;margin:0 0 2px}}h2{{font-size:14px;margin:14px 0 6px;border-bottom:2px solid #e07a5f;padding-bottom:2px}}
.g{{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}}.c{{border:1px solid #ccc;border-radius:8px;padding:6px 10px}}
.t{{color:#555;font-size:11px}}.v{{font-size:22px;font-weight:700}}.a{{font-size:11px;min-height:14px}}
table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #ccc;padding:3px 6px;text-align:left}}th{{background:#f0eee9}}
.n{{color:#666;font-size:11px}}@media print{{body{{padding:0}}}}</style></head><body>
<h1>Attendance report · {esc(rep['title'])}</h1><div class='n'>Xin Heng Chang · {k['Workers']} workers · {k['Good workers']} good · {k['Watch workers']} watch · {k['Bad workers']} bad{(' · compared with ' + esc(prev['title'])) if prev else ''}</div>
<h2>Key numbers</h2><div class='g'>{cards}</div>
<h2>By department</h2>{table(dep)}
<h2>Workers to talk to (worst 10)</h2>{table(w)}
<h2>Flagged in more than one month</h2>{off}
<p class='n'>Late/early = minutes written in the attendance Excel. Missed scans = letters + days with 1-2 punches. Days of a fingerprint problem never count. Print this page (Ctrl+P) and choose "Save as PDF".</p>
</body></html>"""


# ------------------------------------------------------------------ the page
def render(WL, months, db, work):
    import streamlit as st
    st.title("📊 Monthly report for the boss")
    st.caption("Late / early, absent, missed punches, paper different from the fingerprint, leave and OT, for the whole factory and for each department, "
               "with last month next to it. You can download an Excel file (one sheet per topic) or a one-page report to print. Nothing here changes your Excel.")
    if not months:
        st.info("No attendance Excel with a readable month was found. Upload it in Daily work, Step 1.")
        return
    rules = SC.load_rules(os.path.join(work, "score_rules.json"))
    ym_list = sorted(months, reverse=True)
    ym = st.selectbox("📅 Month", ym_list, key="mr_month", format_func=lambda k_: f"{WL.EN_MONTHS[k_[1] - 1]} {k_[0]}  ·  {os.path.basename(months[k_])}")
    cache = st.session_state.setdefault("mr_cache", {})
    letters, cases, outages = db.letters_list(), db.case_list(), db.outages_list(500)

    def get(k_):
        sig = (months[k_], os.path.getmtime(months[k_]), len(letters), len(cases), len(outages), tuple(sorted(rules.items())))
        if sig not in cache:
            with st.spinner(f"Reading {os.path.basename(months[k_])} (about 10 seconds)..."):
                loaded = WL.load(months[k_])
                cache[sig] = build(loaded, k_[0], k_[1], letters, cases, outages, rules)
        return cache[sig]
    rep = get(ym)
    before = [k_ for k_ in ym_list if k_ < ym]
    prev = get(before[0]) if before and (before[0][0] * 12 + before[0][1]) == (ym[0] * 12 + ym[1] - 1) else None
    if prev is None:
        st.caption("No Excel for the month before, so there is no comparison.")
    k = rep["kpi"]

    # ---- key numbers with arrows
    shown = ["Attendance rate %", "Absent days (A)", "Late / early incidents", "Late / early minutes", "Missed scans", "Paper ≠ fingerprint", "Leave days", "OT hours (all kinds)"]
    for row in (shown[:4], shown[4:]):
        cols = st.columns(4)
        for c_, n_ in zip(cols, row):
            b = prev["kpi"].get(n_) if prev else None
            a_, col = arrow(n_, k[n_], b)
            c_.metric(n_, f"{k[n_]:g}", delta=(None if not a_ else a_.replace("▲", "+").replace("▼", "-").replace("＝", "0")),
                      delta_color=("inverse" if n_ in WORSE_UP else "normal"), help="Compared with last month" if prev else None)
    g1, g2, g3, g4 = st.columns(4)
    g1.metric("Workers", k["Workers"])
    g2.metric("🟢 Good", k["Good workers"])
    g3.metric("🟡 Watch", k["Watch workers"])
    g4.metric("🔴 Bad", k["Bad workers"])
    st.caption(f"Bad / Watch / Good use your rules from the Scorecard page (Bad at {rules['missed']} missed scans, {rules['late']} late minutes, {rules['absent']} absences, {rules['ot']} OT questions).")

    # ---- departments
    st.subheader("Departments")
    dep = rep["dept"]
    st.dataframe(dep, hide_index=True, width="stretch")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Late / early minutes**")
        st.bar_chart(dep.set_index("Department")[["Late/early min"]], color=["#2563eb"])
    with c2:
        st.markdown("**Absences (A) and missed scans**")
        st.bar_chart(dep.set_index("Department")[["Absent (A)", "Missed scans"]], color=["#dc2626", "#f59e0b"])
    with c3:
        st.markdown("**Good / Watch / Bad workers**")
        st.bar_chart(dep.set_index("Department")[["Good", "Watch", "Bad"]], color=["#16a34a", "#eab308", "#dc2626"])
    if len(rep["trend"]):
        st.subheader("Day by day")
        tr = rep["trend"].copy()
        tr["Day"] = pd.to_datetime(tr["Day"]).dt.strftime("%d %b")
        st.bar_chart(tr.set_index("Day"))
        st.caption("How many lines each day: late / early, absent, missed scans and paper ≠ fingerprint. A high bar = a day to look at.")

    # ---- topics
    st.subheader("Who, for each topic")
    names = [("late", "⏰ Late or early"), ("absent", "🚫 Absent (A)"), ("missed", "👆 Missed punch"), ("paper", "📄 Paper ≠ fingerprint (OT)"),
             ("nofp", "On paper but no fingerprint"), ("notpaper", "Fingerprint but not on paper")]
    tabs = st.tabs([f"{t} ({len(rep['topics'][key])})" for key, t in names])
    for tb, (key, t) in zip(tabs, names):
        with tb:
            d = rep["topics"][key]
            if not len(d):
                st.success("Nothing this month.")
                continue
            if key in ("late", "absent", "missed", "paper"):
                per = d.groupby(["Worker", "Name", "Department"]).size().reset_index(name="Times").sort_values("Times", ascending=False)
                st.markdown("**Top workers**")
                st.dataframe(per.head(15), hide_index=True, width="stretch")
            st.markdown("**Every line**")
            st.dataframe(d, hide_index=True, width="stretch", height=min(80 + 36 * len(d), 450))

    # ---- worst and repeat
    st.subheader("Workers to talk to")
    st.dataframe(SC.style(worst(rep, 15)), hide_index=True, width="stretch")
    reps_3 = [rep] + ([prev] if prev else [])
    earlier = [k_ for k_ in ym_list if k_ < ym and (prev is None or k_ != before[0])][:1]
    if earlier:
        reps_3.append(get(earlier[0]))
    off = repeat_offenders(reps_3)
    st.markdown("**Flagged (Bad or Watch) in 2 or more of the last months**")
    if len(off):
        st.dataframe(off, hide_index=True, width="stretch")
    else:
        st.caption("None, or there is only one month of data.")

    # ---- files
    st.subheader("Give it to the boss")
    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Excel file (one sheet per topic)", to_xlsx(rep, prev, off), file_name=f"attendance_report_{ym[0]}-{ym[1]:02d}.xlsx", key="mr_xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("🖨️ One page to print (open it, press Ctrl+P, Save as PDF)", to_html(rep, prev, off).encode("utf-8"), file_name=f"attendance_report_{ym[0]}-{ym[1]:02d}.html",
                       mime="text/html", key="mr_html")
