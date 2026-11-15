"""Page 'Chinese workers' (used by app.py)."""
import os
from datetime import date, datetime

import pandas as pd
import streamlit as st

import chinese_attendance as C
import attendance_agent as A


def render(work_dir):
    folder = os.path.join(work_dir, "chinese")
    os.makedirs(folder, exist_ok=True)
    st.header("🇨🇳 Chinese workers")
    st.caption("Day and night are not separated. Total hours go in the day-hours column (Sunday → column (d), holiday → column (b)). One punch = 10.5. No punch = A. "
               "Anything I am not sure about, I ask you. Nothing is guessed.")

    c1, c2 = st.columns(2)
    xl = c1.file_uploader("1) Chinese attendance Excel (sheets CV / Day shift)", type=["xlsx", "xlsm"], key="cn_xl")
    fpf = c2.file_uploader("2) Chinese fingerprint file", type=["xlsx", "xls", "csv"], key="cn_fp")
    if xl is not None:
        path = os.path.join(folder, xl.name)
        if st.session_state.get("cn_xl_name") != (xl.name, xl.size):
            open(path, "wb").write(xl.getvalue())
            st.session_state["cn_xl_name"] = (xl.name, xl.size)
            st.session_state.pop("cn_plan", None)
        st.session_state["cn_path"] = path
    if fpf is not None:
        fpath = os.path.join(folder, "fp_" + fpf.name)
        if st.session_state.get("cn_fp_name") != (fpf.name, fpf.size):
            open(fpath, "wb").write(fpf.getvalue())
            st.session_state["cn_fp_name"] = (fpf.name, fpf.size)
            st.session_state.pop("cn_plan", None)
        st.session_state["cn_fpath"] = fpath
    path, fpath = st.session_state.get("cn_path"), st.session_state.get("cn_fpath")
    if not path or not fpath or not os.path.exists(path) or not os.path.exists(fpath):
        st.info("Upload both files to start.")
        return

    try:
        fp = C.read_fp(fpath)
        _, info = C.read_book(path)
    except Exception as e:
        st.error(f"I cannot read the files: {e}")
        return
    fmin, fmax = fp["dt"].min().date(), fp["dt"].max().date()
    st.success(f"Excel month: **{info['start']:%d %b %Y} → {info['end']:%d %b %Y}**.  Fingerprint: **{fmin:%d %b %Y} → {fmax:%d %b %Y}**, {len(fp)} punches, {fp['id'].nunique()} people.")
    if fmin < info["start"] or fmax > info["end"]:
        st.warning("The fingerprint is longer than this Excel month. Only days inside the Excel month are filled. "
                   "For the next month use that month's Excel (cells A3 = first day, A2 = last day).")
    lo, hi = max(fmin, info["start"]), min(fmax, info["end"])
    if lo > hi:
        st.error("The fingerprint has no day inside this Excel month. Did you upload the right Excel?")
        return
    rng = st.date_input("Fill these days", value=(lo, hi), min_value=info["start"], max_value=info["end"], key="cn_range")
    if not isinstance(rng, (tuple, list)) or len(rng) != 2:
        st.info("Pick the last day too.")
        return
    d_from, d_to = rng

    # ---- people
    people = C.match_people(fp, info)
    st.subheader("Step 1 · Who is who")
    odd = [p for p in people if p["status"] != "ok"]
    if not odd:
        st.write("✅ Every person in the fingerprint is already in the Excel.")
    rows = []
    for p in odd:
        rows.append({"Fingerprint ID": p["fp_id"], "Name in fingerprint": p["name"], "What I found": {
            "name_diff": "Same ID but a different name in the Excel", "by_name": "Name found in the Excel with another ID", "new": "New worker, not in the Excel"}[p["status"]],
            "Excel ID / name": f"{p['target']}  {p['excel_name']}".strip(), "Do this": "Use this Excel ID" if p["status"] != "new" else "Add as new worker"})
    pdf = pd.DataFrame(rows)
    dec = {}
    if len(pdf):
        st.caption("Please check each line. Choose what I should do. Nothing is saved yet.")
        opts = ["Use this Excel ID", "Add as new worker", "Skip this person"]
        ed = st.data_editor(pdf, hide_index=True, width="stretch", key="cn_people",
                            disabled=["Fingerprint ID", "Name in fingerprint", "What I found", "Excel ID / name"],
                            column_config={"Do this": st.column_config.SelectboxColumn("Do this", options=opts, required=True)})
        for (_, r), p in zip(ed.iterrows(), odd):
            if r["Do this"] == "Skip this person":
                dec[p["fp_id"]] = ""
            elif r["Do this"] == "Add as new worker":
                dec[p["fp_id"]] = C.std_id(p["fp_id"])
            else:
                dec[p["fp_id"]] = p["target"]
    new_ids = {C.std_id(p["fp_id"]) for p in odd if dec.get(p["fp_id"]) == C.std_id(p["fp_id"]) and C.nid(p["fp_id"]) not in {C.nid(i) for i in info["cv"]}}
    if new_ids:
        st.caption("New workers get: ID, English name, entry date = their first punch (or the first day of the month). Group and job title stay empty for you to fill in the CV sheet.")

    # ---- plan
    st.subheader("Step 2 · What will be filled")
    plan = C.make_plan(fp, info, people, dec, d_from, d_to, A.day_type)
    cells, asks = plan["cells"], plan["asks"]
    k1, k2, k3 = st.columns(3)
    k1.metric("Cells I am sure about", len(cells))
    k2.metric("Codes A (no punch)", sum(1 for c in cells if c["kind"] == "code"))
    k3.metric("Questions for you", len(asks))
    with st.expander("See the table (rows = workers, columns = days)", expanded=False):
        st.dataframe(C.grid(plan, info, d_from, d_to), hide_index=True, width="stretch")
        st.caption("❓ = I need your answer below. (Sun) = Sunday column. (Hol) = holiday column. A = no punch.")

    answers = []
    if asks:
        st.subheader("Step 3 · Your answers")
        st.caption("Type the hours in **Your answer** (for example 10.5), or **A**, or leave it empty to skip that day for now. My suggestion is only a hint.")
        ad = pd.DataFrame([{"ID": a["cid"], "Name": a["name"], "Day": a["day"].strftime("%a %d %b"), "Punches": a.get("punches", ""), "Why I ask": a["ask"],
                            "My suggestion": "" if a.get("suggest") is None else str(C._num(a["suggest"])),
                            "Column": {"day": "day hours", "sun": "(d) Sunday", "hol": "(b) holiday", "both": "?"}.get(a["kind"], ""),
                            "Your answer": ""} for a in asks])
        ed2 = st.data_editor(ad, hide_index=True, width="stretch", height=min(60 + 36 * len(ad), 600), key="cn_asks",
                             disabled=[c for c in ad.columns if c != "Your answer"])
        for (i, r), a in zip(ed2.iterrows(), asks):
            v = str(r["Your answer"]).strip()
            if not v:
                continue
            if v.upper() == "A":
                answers.append({"cid": a["cid"], "day": a["day"], "kind": "code", "value": "A"})
                continue
            try:
                num = float(v.replace(",", "."))
            except ValueError:
                st.error(f"{a['cid']} {a['day']:%d %b}: '{v}' is not a number or A.")
                return
            if a["kind"] == "both":
                st.warning(f"{a['cid']} {a['day']:%d %b} is Sunday AND holiday: tell me which column (I will not guess). Row skipped.")
                continue
            answers.append({"cid": a["cid"], "day": a["day"], "kind": a["kind"], "value": num})
        st.write(f"Answered: **{len(answers)}** of {len(asks)}. The others stay empty.")

    # ---- save
    st.subheader("Step 4 · Save to Excel")
    writes = [{k: c[k] for k in ("cid", "day", "kind", "value")} for c in cells] + answers
    nw = []
    for p in odd:
        if dec.get(p["fp_id"]) and dec[p["fp_id"]] in new_ids:
            nw.append({"id": dec[p["fp_id"]], "name": p["name"], "entry": max(p["first"].date(), info["start"])})
    st.write(f"I will write **{len(writes)}** cells and add **{len(nw)}** new worker(s). A cell that already has something is never changed. A backup of the Excel is made first.")
    ok = st.text_input("Type YES to save", key="cn_yes")
    if st.button("💾 Save", type="primary", disabled=ok.strip().upper() != "YES", key="cn_save"):
        try:
            log, n, bad = C.apply_changes(path, writes, nw, path, os.path.join(folder, "backup"))
        except Exception as e:
            st.error(f"Not saved: {e}")
            return
        if bad:
            st.error(f"⚠️ {len(bad)} cells were not saved correctly. First: {bad[0]}. Your backup is in {os.path.join(folder, 'backup')}.")
        else:
            st.success(f"✅ Saved {n} cells. Every cell was read back and checked. Backup: {os.path.join(folder, 'backup')}")
        skipped = [l for l in log if l.startswith("SKIPPED")]
        if skipped:
            st.warning(f"{len(skipped)} cells were not changed because they already had something.")
            with st.expander("Which ones"):
                st.write("\n".join(skipped[:200]))
        with st.expander("Details"):
            st.write("\n".join(log) or "-")
        st.session_state["cn_done"] = path
    if st.session_state.get("cn_done") == path and os.path.exists(path):
        st.download_button("⬇️ Download the Excel", open(path, "rb").read(), file_name=os.path.basename(path), key="cn_dl")
