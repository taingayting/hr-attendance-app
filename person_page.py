"""Page 'One person': fill the days of ONE worker from the fingerprint (same logic as Daily work, no paper), and set resigned / not resigned.
Nothing is written to the Excel without your YES. Every cell is shown Was -> New, a backup is made, every cell is read back, Undo is possible."""
import calendar
import os
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st


def _num(v):
    s = str(v if v is not None else "").strip().replace(",", ".")
    if s in ("", "None", "nan"):
        return None, ""
    try:
        return float(s), ""
    except ValueError:
        return None, f"'{s}' is not a number"


def _workers(att):
    import hr_db as H
    try:
        return H.read_cv_sheet(att)[0]
    except Exception:
        return []


def render(ctx):
    A, db, UK = ctx["A"], ctx["db"], ctx["UK"]
    st.title("🧍 One person / some workers")
    st.caption("For ONE worker: fill many days from the fingerprint (no paper, same rules as Daily work), or say that he resigned / did not resign. "
               "Nothing goes into your Excel before you type YES. You see every cell as Was → New, a backup is made, and Undo can put it back.")
    t1, t2 = st.tabs(["📅 Fill days from the fingerprint", "🚪 Resigned / not resigned"])
    with t1:
        _fill_tab(ctx, A, db, UK)
    with t2:
        _resign_tab(ctx, A, db, UK)


# ------------------------------------------------------------------ fill days
def _fill_tab(ctx, A, db, UK):
    fp_path = ctx["fp_path"]
    if not fp_path or not os.path.exists(fp_path):
        st.info("Give me the fingerprint file first (Daily work, Step 1).")
        return
    try:
        fp = A.read_fingerprint(fp_path)
    except Exception as e:
        st.error(f"I cannot read the fingerprint file: {e}")
        return
    fmin, fmax = fp["dt"].min().date(), fp["dt"].max().date()
    st.success(f"The fingerprint file covers **{fmin:%d %b %Y} → {fmax:%d %b %Y}**. You can choose any days. A day the file does not cover is shown as \"no records\" and nothing is filled for it.")
    first_default = max(fmin, date(fmax.year, fmax.month, 1)) if (fmax.year, fmax.month) != (fmin.year, fmin.month) else fmin
    rng = st.date_input("📅 Which days? (one month at a time)", value=(first_default, fmax), min_value=date(2020, 1, 1), max_value=date(2035, 12, 31), key="op_range")
    if not isinstance(rng, (tuple, list)) or len(rng) != 2:
        st.info("Pick the last day too.")
        return
    d_from, d_to = rng
    if (d_from.year, d_from.month) != (d_to.year, d_to.month):
        st.error("Please choose days of ONE month (the attendance Excel is one month).")
        return
    att = ctx["find_att"](d_from)
    if not att:
        st.error(f"I have no attendance Excel for {d_from:%B %Y}. Upload it in Daily work, Step 1.")
        return
    st.caption(f"Excel used: **{os.path.basename(att)}**")
    ws_ = _workers(att)
    if not ws_:
        st.error("I cannot read the sheet CV of this Excel.")
        return
    label = {w["id"]: f"{w['id']} · {w.get('name_en') or w.get('name_kh') or ''} · {w.get('dept') or ''}" for w in ws_}
    ids = sorted(label)
    _fpids = set(fp["ID Number"])
    in_fp = [i for i in ids if i in _fpids]

    def _all_fp():
        st.session_state["op_workers"] = list(in_fp)

    def _none():
        st.session_state["op_workers"] = []
    st.markdown("**Workers** (choose one or many; type the ID or the name to search)")
    b1, b2, b3 = st.columns([2, 1, 3])
    b1.button(f"➕ Choose all {len(in_fp)} workers of the fingerprint file", key="op_allfp", on_click=_all_fp)
    b2.button("Clear", key="op_none", on_click=_none)
    wids = st.multiselect("Workers", ids, key="op_workers", format_func=lambda i: label[i], label_visibility="collapsed", placeholder="DMO-10015 or part of the name")
    if not wids:
        return
    dept_of = {w["id"]: (w.get("dept") or "") for w in ws_}
    sig = (tuple(sorted(wids)), str(d_from), str(d_to), att, os.path.getmtime(att), os.path.getmtime(fp_path))
    if st.button(f"🔎 Read the fingerprint and build the table ({len(wids)} workers)", type="primary", key="op_build"):
        days = [d_from + timedelta(days=i) for i in range((d_to - d_from).days + 1)]
        empty = os.path.join(ctx["work"], "empty_paper.csv")
        os.makedirs(ctx["work"], exist_ok=True)
        with open(empty, "w", encoding="utf-8-sig") as f:
            f.write("ID,HOURS,OT,CODE,REMARK\n")
        try:
            A.RESIGNED_PREF = db.resigned_prefs()
        except Exception:
            A.RESIGNED_PREF = {}
        try:
            A.SHIFT_PREF = db.shift_prefs()
            A.TIMES_PREF = db.times_prefs()
        except Exception:
            pass
        with st.spinner("Reading the fingerprint day by day (a few seconds per week)..."):
            try:
                A.FILES = {"attendance": att, "fingerprint": fp_path}
                A.CV_PATH = att
                res = A.people_days({w: dept_of.get(w, "") for w in wids}, days, fp_path, att, empty)
            except SystemExit as e:
                st.error(ctx["stop_message"](e))
                return
            except Exception as e:
                ctx["show_problem"]("reading the fingerprint", e)
                return
        st.session_state["op_data"] = {"sig": sig, "days": res}
        st.session_state.pop("op_plan", None)
    data = st.session_state.get("op_data")
    if not data or data["sig"] != sig:
        st.info("Choose the workers and the days, then press the blue button.")
        return
    saved = st.session_state.setdefault("op_user", {})            # what you typed stays, also when you build the table again
    rows = []
    for d in sorted(data["days"], key=lambda x: (x["wid"], x["day"])):
        u = saved.get((d["wid"], str(d["day"])), {})
        has = d["hours"] is not None or bool(d["code"])
        rows.append({"Write": u.get("Write", has), "Worker": d["wid"], "Name": (label[d["wid"]].split(" · ")[1] if d["wid"] in label else ""), "Date": d["day"].strftime("%a %d %b"),
                     "Type": d["type"].capitalize(), "Shift": d["shift"], "Punches": d["punches"], "What I found": d["found"] + (" | " + d["note"] if d["note"] else ""),
                     "Fill hours": u.get("Fill hours", "" if d["hours"] is None else f"{d['hours']:g}"),
                     "Fill OT": u.get("Fill OT", "" if d["ot"] is None else f"{d['ot']:g}"),
                     "Code": u.get("Code", d["code"]), "Late/Early (min)": u.get("Late/Early (min)", "" if not d["le"] else str(int(d["le"])))})
    dsorted = sorted(data["days"], key=lambda x: (x["wid"], x["day"]))
    df = pd.DataFrame(rows)
    ask_n = sum(1 for d in dsorted if d["ask"])
    st.subheader(f"{len(wids)} worker(s), {len(df)} rows")
    st.write(f"Rows that need YOUR answer ({ask_n}): type the hours, OT or code and tick Write. Rows I am sure about are ticked already. "
             f"A row with a note (resigned, no records...) stays unticked.")
    ed = st.data_editor(df, hide_index=True, width="stretch", height=min(90 + 36 * len(df), 700), key=f"op_ed_{abs(hash(sig)) % 10**8}",
                        disabled=["Worker", "Name", "Date", "Type", "Shift", "Punches", "What I found"],
                        column_config={"Write": st.column_config.CheckboxColumn(width=60), "Fill hours": st.column_config.TextColumn(width=80),
                                       "Fill OT": st.column_config.TextColumn(width=70), "Code": st.column_config.SelectboxColumn(options=[""] + A.CODES, width=80),
                                       "Late/Early (min)": st.column_config.TextColumn(width=90), "What I found": st.column_config.TextColumn(width="large")})
    for (i, e), d in zip(ed.iterrows(), dsorted):                      # remember every change
        saved[(d["wid"], str(d["day"]))] = {k: e[k] for k in ("Write", "Fill hours", "Fill OT", "Code", "Late/Early (min)")}

    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        ed.drop(columns=["Write"]).to_excel(xw, sheet_name="Workers", index=False)
    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Download this table (Excel file)", buf.getvalue(), file_name=f"workers_{d_from:%Y-%m-%d}_{d_to:%Y-%m-%d}.xlsx", key="op_dl_tbl",
                       help="The table as you see it now, with the numbers you typed.")
    d2.download_button("⬇️ Download the monthly attendance Excel", ctx["lazy_file"](att), file_name=os.path.basename(att), key="op_dl_att", on_click="ignore",
                       help="The whole month Excel (with everything saved so far).")
    if st.button("🔎 Check what would change in the Excel", key="op_check"):
        items, errs = [], []
        for (i, e), d in zip(ed.iterrows(), dsorted):
            if not e["Write"]:
                continue
            h, e1 = _num(e["Fill hours"])
            o, e2 = _num(e["Fill OT"])
            le, e3 = _num(e["Late/Early (min)"])
            for m in (e1, e2, e3):
                if m:
                    errs.append(f"{d['wid']} {d['day']:%d %b}: {m}")
            code = str(e["Code"] or "").strip()
            if h is None and not code:
                errs.append(f"{d['wid']} {d['day']:%d %b}: ticked Write but no hours and no code.")
                continue
            items.append({"wid": d["wid"], "day": d["day"], "hours": h, "ot": o, "code": code, "night": d["night"], "le": int(le) if le else None})
        if errs:
            for m in errs[:30]:
                st.error(m)
        elif not items:
            st.info("Nothing is ticked.")
        else:
            try:
                with st.spinner("Preparing the Excel (about 10 seconds)..."):
                    A.CV_PATH = att
                    A.EXTRA_SKIP = set()
                    path_, wb_, log_, ch_ = A.plan_people(items, fp_path, att)
                st.session_state["op_plan"] = {"sig": sig, "path": path_, "wb": wb_, "log": log_, "changes": ch_, "mtime": os.path.getmtime(path_)}
            except SystemExit as e:
                st.error(ctx["stop_message"](e))
            except Exception as e:
                ctx["show_problem"]("checking the Excel", e)
    pl = st.session_state.get("op_plan")
    if st.session_state.get("op_flash"):
        st.success(st.session_state.pop("op_flash"))
    if not pl or pl["sig"] != sig:
        return
    if pl["mtime"] != os.path.getmtime(pl["path"]):
        st.session_state.pop("op_plan", None)
        st.info("The Excel changed meanwhile. Press the check button again.")
        return
    for l in pl["log"]:
        if "INFO" not in l:
            st.warning(l)
    ch = pl["changes"]
    if not ch:
        st.info("Nothing to change: the Excel already has these numbers.")
        return
    cdf = pd.DataFrame([{"Write": True, "Worker": w_, "Cell": c_, "Was": "" if o_ in (None, "") else o_, "Will be": "(empty)" if n_ in (None, "") else n_} for w_, c_, o_, n_ in ch])
    st.write(f"**{len(cdf)} cells** would change for **{cdf['Worker'].nunique()} worker(s)**.")
    ed2 = st.data_editor(cdf, hide_index=True, width="stretch", disabled=["Worker", "Cell", "Was", "Will be"], key=f"op_ch_{len(cdf)}_{abs(hash(sig)) % 10**6}",
                         height=min(80 + 36 * len(cdf), 500))
    yes = st.text_input("Happy with this? Type YES (capital letters) and press the button.", key="op_yes")
    if st.button("✅ Save these changes into the Excel", type="primary", key="op_save"):
        if yes.strip() != "YES":
            st.error("Not saved. Please type YES in capital letters.")
            return
        keep = list(ed2["Write"])
        ws2 = pl["wb"][A.SHEET]
        for (w_, c_, o_, n_), k_ in zip(ch, keep):
            if not k_:
                ws2[c_].value = o_
        chosen = [c for c, k_ in zip(ch, keep) if k_]
        if not chosen:
            st.info("You unticked everything. Nothing was saved.")
            return
        A.CHANGES[:] = chosen
        UK.snapshot(db.DB, f"One person: {len(wids)} workers", excel_path=pl["path"], force=True)
        ok, bad = A.save_changes(pl["path"], pl["wb"], d_from, "one person")
        st.session_state.pop("op_plan", None)
        if ok:
            for k_ in ("wl_cache", "sc_load", "excel_now", "hist_plan", "plan"):
                st.session_state.pop(k_, None)
            st.session_state["op_flash"] = f"✅ Changed {len(chosen)} cells for {len({c[0] for c in chosen})} worker(s). I read every cell back and they are correct. A backup was made first. Undo can put it back."
            ctx["sync"]()
            st.rerun()
        else:
            st.error(f"Problem - these cells are not as expected: {bad}. Please use Undo or restore the backup.")


# ------------------------------------------------------------------ resign
def _resign_tab(ctx, A, db, UK):
    import worker_list as WL
    att = ctx["att_path"]
    if not att:
        st.info("Give me the attendance Excel first (Daily work, Step 1).")
        return
    ws_ = _workers(att)
    label = {w["id"]: f"{w['id']} · {w.get('name_en') or w.get('name_kh') or ''} · {w.get('dept') or ''}" for w in ws_}
    wid = st.selectbox("🔎 Worker (type the ID or the name to search)", sorted(label), key="rs_worker", format_func=lambda i: label[i],
                       index=None, placeholder="DMO-10015 or part of the name")
    if not wid:
        st.caption("Resigned workers get nothing filled from their resign date on. Use this when a worker said he resigned and then stayed.")
        return
    mem = db.resigned_prefs().get(wid)
    try:
        cvd = A._cv_dates(att).get(wid, (None, None))[1]
    except SystemExit:
        cvd = None
    book = next((w.get("stop_date") for w in ws_ if w["id"] == wid), "")
    c1, c2, c3 = st.columns(3)
    c1.metric("App memory: resigned since", str(mem) if mem else "not resigned")
    c2.metric("Excel sheet CV: resign date", "-" if cvd in (None, "") else str(cvd)[:10])
    c3.metric("Worker book", str(book) if book else "-")
    mode = st.radio("What do you want?", ["He resigned (set the date)", "He did NOT resign (take the resign away)"], key=f"rs_mode_{wid}", horizontal=True, index=None)
    if not mode:
        st.caption("Choose one. Nothing is changed before you type YES.")
        return
    new_date = None
    if mode.startswith("He resigned"):
        try:
            dflt = date.fromisoformat(str(mem)[:10]) if mem else (date.fromisoformat(str(cvd)[:10]) if cvd else date.today())
        except ValueError:
            dflt = date.today()
        new_date = st.date_input("Resign date (from this day nothing is filled for him)", value=dflt, key=f"rs_date_{wid}")
    also_excel = st.checkbox("Also change the resign date in the Excel (sheet CV)", value=True, key=f"rs_xl_{wid}")
    lines = []
    old_x = "" if cvd in (None, "") else str(cvd)[:10]
    new_x = "" if new_date is None else str(new_date)
    lines.append({"Where": "App memory", "Was": str(mem or ""), "Will be": new_x})
    if also_excel:
        lines.append({"Where": "Excel sheet CV (resign date)", "Was": old_x, "Will be": new_x})
    st.dataframe(pd.DataFrame(lines), hide_index=True, width="stretch")
    yes = st.text_input("Type YES (capital letters) to save", key=f"rs_yes_{wid}")
    if st.button("💾 Save", type="primary", key=f"rs_save_{wid}"):
        if yes.strip() != "YES":
            st.error("Not saved. Please type YES in capital letters.")
            return
        try:
            if new_date is None:
                db.resigned_forget(wid)
            else:
                db.resigned_set(wid, str(new_date))
            msg = "App memory saved."
            if also_excel and old_x != new_x:
                UK.snapshot(db.DB, f"Resign {wid}", excel_path=att, force=True)
                n, probs, bad = WL.apply_edits(att, [{"type": "cv", "wid": wid, "col": 19,
                                                      "new": (datetime.combine(new_date, datetime.min.time()) if new_date else None)}], os.path.join(ctx["work"], "backup"))
                if probs or bad:
                    st.error(f"Excel: {probs or bad}. The app memory was saved, the Excel was not changed correctly. Use Undo or the backup.")
                    return
                msg += f" Excel sheet CV changed and read back. A backup was made."
            for k_ in ("wl_cache", "sc_load", "excel_now", "hist_plan", "plan", "dept_map"):
                st.session_state.pop(k_, None)
            ctx["sync"]()
            st.success(msg + (" Now open the tab 'Fill days from the fingerprint' to fill the days he was skipped." if new_date is None else ""))
        except Exception as e:
            ctx["show_problem"]("saving the resign", e)
