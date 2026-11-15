"""Worker scorecard (Good / Watch / Bad) per worker and per department, per month.
Numbers come from two places:
  * the attendance Excel of that month: late/early minutes, code A (absent), leave codes
  * the app memory: missed scans (letters + '1-2 punches' cases), OT questions, no-fingerprint cases
Days of a fingerprint problem (outage) never count against a worker. Nothing here writes to the Excel."""
import json
import math
import os
from datetime import date, timedelta

import pandas as pd

LEAVE = {"HP": 1, "AL1": 1, "AL2": 1, "MC": 1, "PC": 1, "M": 1, "NP": 1, "HW": 1, "HH": 0.5}
DEFAULT_RULES = {"missed": 3, "late": 60, "absent": 2, "ot": 3}      # RED at or above; WATCH (yellow) at half of it, at least 1
LABEL = {"missed": "missed scans", "late": "late/early minutes", "absent": "absences (A)", "ot": "OT questions"}


def load_rules(path):
    try:
        with open(path, encoding="utf-8") as f:
            r = json.load(f)
        return {k: max(1, int(r.get(k, v))) for k, v in DEFAULT_RULES.items()}
    except Exception:
        return dict(DEFAULT_RULES)


def save_rules(path, rules):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(rules, f)
    except OSError:
        pass


def watch_at(red):
    return max(1, math.ceil(red / 2))


def _outage_days(outages, ym):
    """[(day_from, day_to, DEPT or '')] as date objects."""
    out = []
    for o in outages:
        try:
            a = date.fromisoformat(str(o["day"]))
            b = date.fromisoformat(str(o.get("to_day") or o["day"]))
        except ValueError:
            continue
        out.append((a, b, str(o.get("dept") or "").strip().upper()))
    return out


def _in_outage(outs, d, dept):
    return any(a <= d <= b and (not od or od == str(dept).strip().upper()) for a, b, od in outs)


def compute(loaded, year, month, letters, cases, outages, rules):
    """-> (DataFrame one row per worker, details dict worker_id -> list of (day, what))"""
    ym = f"{year}-{month:02d}"
    outs = _outage_days(outages, ym)
    dept_of = {w["id"]: w["dept"] for w in loaded["workers"]}
    miss, otq, nofp, det = {}, {}, {}, {}

    def add_det(wid, d, what):
        det.setdefault(wid, []).append((d, what))

    for L in letters:
        d = str(L["day"])
        if not d.startswith(ym) or L["worker_id"] not in dept_of:
            continue
        if _in_outage(outs, date.fromisoformat(d), dept_of[L["worker_id"]]):
            continue
        miss.setdefault(L["worker_id"], set()).add(d)
        add_det(L["worker_id"], d, "Missed scan (letter): " + str(L.get("problem") or ""))
    for c in cases:
        d, wid = str(c["day"]), c["worker_id"]
        if not d.startswith(ym) or wid not in dept_of:
            continue
        if _in_outage(outs, date.fromisoformat(d), dept_of[wid]):
            continue
        ans = (" — asked: " + str(c["answer"])) if c.get("answer") else (" — asked on " + str(c["asked_on"]) if c.get("asked_on") else " — not asked yet")
        if c["kind"] == "1-2 punches":
            if d not in miss.get(wid, set()):
                add_det(wid, d, "Missed scan (1-2 punches)" + ans)
            miss.setdefault(wid, set()).add(d)
        elif c["kind"] == "OT different":
            otq.setdefault(wid, set()).add(d)
            add_det(wid, d, f"OT different (paper {c.get('paper_ot', '')} / fingerprint {c.get('finger_ot', '')})" + ans)
        elif c["kind"] in ("No fingerprint", "Not on paper"):
            nofp.setdefault(wid, set()).add(d)
            add_det(wid, d, c["kind"] + ans)
    rows = []
    for w in loaded["workers"]:
        wid = w["id"]
        per = loaded["data"].get(wid, {})
        late_min = late_days = absent = 0
        leave = 0.0
        for dn, x in per.items():
            try:
                dd = date(year, month, int(dn))
            except ValueError:
                continue
            code = str(x.get("code") or "").strip()
            if _in_outage(outs, dd, w["dept"]):
                continue
            if x.get("late") and x["late"] > 0:
                late_min += x["late"]
                late_days += 1
                add_det(wid, dd.isoformat(), f"Late/early {x['late']:g} min")
            if code.upper() == "A":
                absent += 1
                add_det(wid, dd.isoformat(), "Absent (A)")
            elif code.upper() in LEAVE:
                leave += LEAVE[code.upper()]
        val = {"missed": len(miss.get(wid, ())), "late": late_min, "absent": absent, "ot": len(otq.get(wid, ()))}
        red = [f"{val[k]:g} {LABEL[k]}" for k in rules if val[k] >= rules[k]]
        yel = [f"{val[k]:g} {LABEL[k]}" for k in rules if watch_at(rules[k]) <= val[k] < rules[k]]
        lvl = "Bad" if red else "Watch" if yel else "Good"
        rows.append({"ID": wid, "Name": w.get("kh") or w.get("en") or "", "English": w.get("en") or "", "Department": w["dept"], "Result": lvl,
                     "Why": "; ".join(red or yel), "Missed scans": val["missed"], "Late/early min": val["late"], "Late days": late_days, "Absent (A)": absent,
                     "Leave days": leave, "OT questions": val["ot"], "No fp / not on paper": len(nofp.get(wid, ())), "_score": len(red) * 100 + len(yel) * 10 + val["missed"] + val["absent"]})
    return pd.DataFrame(rows), det


def by_department(df):
    g = []
    for dept, x in df.groupby("Department"):
        n = len(x)
        g.append({"Department": dept, "Workers": n, "Bad": int((x["Result"] == "Bad").sum()), "Watch": int((x["Result"] == "Watch").sum()),
                  "Good": int((x["Result"] == "Good").sum()), "% Good": round(100 * (x["Result"] == "Good").sum() / n),
                  "Missed scans": int(x["Missed scans"].sum()), "Late/early min": float(x["Late/early min"].sum()), "Absent (A)": int(x["Absent (A)"].sum()),
                  "OT questions": int(x["OT questions"].sum())})
    return pd.DataFrame(g).sort_values(["% Good", "Department"]).reset_index(drop=True)


COLORS = {"Bad": "background-color:#fecaca;color:#7f1d1d", "Watch": "background-color:#fef08a;color:#713f12", "Good": "background-color:#bbf7d0;color:#14532d"}


def style(df):
    return df.style.map(lambda v: COLORS.get(v, ""), subset=["Result"])


# ============================================================ the page
def render(WL, months, db, work):
    import io
    import streamlit as st
    st.title("🏅 Worker scorecard")
    st.caption("Good / Watch / Bad for every worker and department, one month at a time. Late minutes and absences come from your attendance Excel. "
               "Missed scans (letters and 1-2 punch cases) and OT questions come from the app memory. Days of a fingerprint problem never count. Nothing here changes the Excel.")
    if not months:
        st.info("No attendance Excel with a readable month was found. Upload it in Daily work, Step 1.")
        return
    rules_path = os.path.join(work, "score_rules.json")
    rules = load_rules(rules_path)
    ym_list = sorted(months, reverse=True)
    c1, c2 = st.columns([1, 1])
    ym = c1.selectbox("📅 Month", ym_list, key="sc_month", format_func=lambda k: f"{WL.EN_MONTHS[k[1] - 1]} {k[0]}  ·  {os.path.basename(months[k])}")
    cache = st.session_state.setdefault("sc_load", {})

    def get_loaded(k):
        sig = (months[k], os.path.getmtime(months[k]))
        if sig not in cache:
            with st.spinner(f"Reading {os.path.basename(months[k])} (about 10 seconds)..."):
                cache[sig] = WL.load(months[k])
        return cache[sig]

    loaded = get_loaded(ym)
    letters, cases, outages = db.letters_list(), db.case_list(), db.outages_list(500)
    with st.expander("⚙️ The rules (change the numbers if you want)"):
        cols = st.columns(4)
        new = {k: int(cols[i].number_input(f"Bad at {LABEL[k]} ≥", 1, 999, rules[k], key=f"sc_r_{k}")) for i, k in enumerate(DEFAULT_RULES)}
        st.caption("Watch (yellow) starts at half of each number. A missed scan counts fully, with or without a letter.")
        if new != rules:
            save_rules(rules_path, new)
            rules = new
    df, det = compute(loaded, ym[0], ym[1], letters, cases, outages, rules)
    depts = ["All departments"] + sorted(df["Department"].unique())
    pick = c2.selectbox("Department", depts, key="sc_dept")
    view = df if pick == "All departments" else df[df["Department"] == pick]
    m = st.columns(4)
    m[0].metric("Workers", len(view))
    m[1].metric("🔴 Bad", int((view["Result"] == "Bad").sum()))
    m[2].metric("🟡 Watch", int((view["Result"] == "Watch").sum()))
    m[3].metric("🟢 Good", int((view["Result"] == "Good").sum()))

    st.subheader("Departments")
    dep = by_department(df)
    st.dataframe(dep, hide_index=True, width="stretch")
    st.subheader("Workers (worst first)")
    show = view.sort_values(["_score", "ID"], ascending=[False, True]).drop(columns=["_score"])
    st.dataframe(style(show), hide_index=True, width="stretch", height=min(80 + 36 * len(show), 650))

    st.subheader("One worker")
    ids = list(show["ID"])
    if ids:
        wid = st.selectbox("Worker", ids, key="sc_worker", format_func=lambda i: f"{i}  {view.loc[view['ID'] == i, 'Name'].iloc[0]}  ({view.loc[view['ID'] == i, 'Result'].iloc[0]})")
        row = df[df["ID"] == wid].iloc[0]
        st.write(f"**{row['English'] or row['Name']}** · {row['Department']} · **{row['Result']}**" + (f" — {row['Why']}" if row["Why"] else ""))
        try:
            import notes_ui
            notes_ui.worker_history(wid)
        except Exception:
            pass
        items = sorted(det.get(wid, []))
        if items:
            st.dataframe(pd.DataFrame(items, columns=["Day", "What happened"]), hide_index=True, width="stretch")
        else:
            st.success("Nothing to report this month.")
        if st.button("Show this worker in every month I have", key="sc_trend"):
            tr = []
            for k in sorted(months):
                d2, _ = compute(get_loaded(k), k[0], k[1], letters, cases, outages, rules)
                r2 = d2[d2["ID"] == wid]
                if len(r2):
                    r2 = r2.iloc[0]
                    tr.append({"Month": f"{WL.EN_MONTHS[k[1] - 1]} {k[0]}", "Result": r2["Result"], "Missed scans": r2["Missed scans"], "Late/early min": r2["Late/early min"],
                               "Absent (A)": r2["Absent (A)"], "Leave days": r2["Leave days"], "OT questions": r2["OT questions"]})
            st.dataframe(style(pd.DataFrame(tr)) if tr else pd.DataFrame(), hide_index=True, width="stretch")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        dep.to_excel(xw, sheet_name="Departments", index=False)
        df.sort_values(["_score"], ascending=False).drop(columns=["_score"]).head(10).to_excel(xw, sheet_name="Worst 10", index=False)
        for d_, x in df.groupby("Department"):
            x.sort_values("_score", ascending=False).drop(columns=["_score"]).to_excel(xw, sheet_name=(str(d_)[:30].replace("/", "-") or "dept"), index=False)
    st.download_button("⬇️ Save this month as an Excel report", buf.getvalue(), file_name=f"scorecard_{ym[0]}-{ym[1]:02d}.xlsx", key="sc_dl")
