"""HR pages for the website: Workers list + one worker profile (stage 2)."""
import html
import io
import os
import zipfile
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

import attendance_db as db
import hr_db as H


def _d(s):
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def _iso(d):
    return d.isoformat() if d else ""


def _csv(df):
    """Excel runs text that starts with = + - @ as a formula: put a quote in front."""
    df = df.copy()
    for c in df.columns:
        df[c] = df[c].map(lambda v: "'" + v if isinstance(v, str) and v[:1] in ("=", "+", "-", "@", "\t", "\r") else v)
    return df.to_csv(index=False).encode("utf-8-sig")


MIN_DATE, MAX_DATE = date(1940, 1, 1), date(2100, 12, 31)

HR_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Figtree:wght@400;500;600;700&family=Fraunces:opsz,wght@9..144,600&display=swap');
[data-testid="stAppViewContainer"], [data-testid="stApp"] {background: #F7F4EE;}
[data-testid="stHeader"] {background: transparent;}
[data-testid="stSidebar"] {background: #EFEAE0;}
html, body, [data-testid="stAppViewContainer"] p, [data-testid="stAppViewContainer"] label {font-family: 'Figtree', system-ui, sans-serif;}
[data-testid="stAppViewContainer"] h1, [data-testid="stAppViewContainer"] h2, [data-testid="stAppViewContainer"] h3 {font-family: 'Fraunces', Georgia, serif !important; font-weight: 600; color: #26302C; letter-spacing: -0.01em;}
[data-testid="stAppViewContainer"] h1 {font-size: 2.6rem; line-height: 1.1;}
[data-testid="stAppViewContainer"] h4 {font-family: 'Figtree', system-ui, sans-serif !important; font-weight: 700; color: #26302C;}
[data-testid="stMetric"] {background: #FFFFFF; border-radius: 18px; padding: 18px 22px; box-shadow: 0 1px 2px rgba(38,48,44,0.08);}
[data-testid="stMetricLabel"] p {color: #5E6B65; font-size: 0.95rem; font-weight: 500;}
[data-testid="stMetricValue"] {font-family: 'Fraunces', Georgia, serif !important; font-size: 2.6rem; font-weight: 600; color: #26302C;}
[data-testid="stExpander"] {background: #FFFFFF; border: 0 !important; border-radius: 18px; box-shadow: 0 1px 2px rgba(38,48,44,0.08);}
[data-testid="stExpander"] summary {font-weight: 600;}
.st-key-hr_head {background: #FFFFFF !important; border: 0 !important; border-radius: 20px !important; box-shadow: 0 1px 2px rgba(38,48,44,0.08); padding: 18px 20px !important;}
[data-testid="stDataFrame"] {background: #FFFFFF; border-radius: 18px; padding: 8px; box-shadow: 0 1px 2px rgba(38,48,44,0.08);}
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {border-radius: 12px; font-weight: 600; border: 1px solid #D5DCD6; min-height: 44px;}
.stButton > button[kind="primary"], [data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"] {background: #1F6F5F; border-color: #1F6F5F; color: #FFFFFF;}
.stButton > button[kind="primary"]:hover, [data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"]:hover {background: #134B40; border-color: #134B40;}
[data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea, [data-baseweb="select"] > div, [data-testid="stDateInput"] input {border-radius: 12px; background: #FBFAF7;}
[role="tablist"] {gap: 8px; border-bottom: 0 !important; box-shadow: none !important;}
[role="tablist"] > div:not([role="tab"]):not([data-testid="stTab"]) {display: none !important;}
[data-testid="stTab"] {border-radius: 999px !important; padding: 8px 18px !important; background: #FFFFFF !important; border: 1px solid #D5DCD6 !important; height: auto !important; color: #26302C !important;}
[data-testid="stTab"][data-selected="true"] {background: #1F6F5F !important; border-color: #1F6F5F !important;}
[data-testid="stTab"][data-selected="true"] p, [data-testid="stTab"][data-selected="true"] span {color: #FFFFFF !important;}
[data-testid="stSidebar"] [data-testid="stRadio"] label {border-radius: 999px;}
.hr-pill {display: inline-block; font-size: 0.8rem; font-weight: 700; padding: 5px 12px; border-radius: 999px; margin-right: 8px;}
.hr-green {background: #DCEEE8; color: #1B5446;} .hr-red {background: #F6D9D4; color: #8E2D23;} .hr-amber {background: #FFE3A3; color: #5A3E00;}
.hr-sub {color: #5E6B65; font-size: 0.95rem;}
.hr-avatar {width: 120px; height: 120px; border-radius: 50%; background: #DCEEE8; color: #1B5446; display: flex; align-items: center; justify-content: center; font-family: 'Fraunces', Georgia, serif; font-size: 42px; font-weight: 600;}
</style>
"""


def _theme():
    st.markdown(HR_CSS, unsafe_allow_html=True)



# ================================================================== the list
def render_workers(att_path=None):
    _theme()
    if st.session_state.get("hr_open"):
        render_profile(st.session_state["hr_open"])
        return
    st.markdown(f"<div class='hr-sub'>{date.today():%A, %d %B %Y}</div>", unsafe_allow_html=True)
    st.title("Workers")
    ws = H.list_workers()
    flash = st.session_state.pop("hr_flash", "")
    if flash:
        st.success(flash)

    working = [w for w in ws if w["status"] == "Working"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Working now", len(working))
    c2.metric("Women", sum(w["sex"] == "F" for w in working))
    c3.metric("Men", sum(w["sex"] == "M" for w in working))
    c4.metric("With documents missing", sum(w["docs_missing"] > 0 for w in working))

    with st.expander("⬇️ Import the workers from my Excel (sheet CV)", expanded=not ws):
        st.caption("Reads the sheet CV of your Step 1 attendance Excel. The Excel is never changed. "
                   "Workers who are already in the worker book are never changed by an import: only new IDs are added.")
        if not att_path:
            st.info("Put your attendance Excel in Step 1 of Daily work first.")
        else:
            st.write(f"File: **{att_path.split('/')[-1].split(chr(92))[-1]}**")
            if st.button("1. Read the CV sheet (nothing is saved yet)", key="hr_read_cv"):
                try:
                    items, warns = H.read_cv_sheet(att_path)
                    st.session_state["hr_import"] = {"items": items, "warns": warns, "prev": H.import_preview(items)}
                except Exception as e:
                    st.session_state.pop("hr_import", None)
                    st.error(f"I could not read the Excel ({type(e).__name__}). Close it in Excel and try again.")
            imp = st.session_state.get("hr_import")
            if imp:
                p = imp["prev"]
                st.info(f"The sheet has **{p['total']}** workers. **{p['new']}** are new and **{p['existing']}** are already in the worker book.")
                if imp["warns"]:
                    with st.expander(f"⚠️ {len(imp['warns'])} thing(s) I could not read (these cells are left empty, you can fix them in the profile)"):
                        for x in imp["warns"][:200]:
                            st.markdown("- " + x)
                if p["new"] and st.button(f"2. Add the {p['new']} new workers to the worker book", type="primary", key="hr_do_import"):
                    n = H.import_commit(imp["items"])
                    st.session_state.pop("hr_import", None)
                    st.session_state["hr_flash"] = f"Added {n} workers to the worker book."
                    st.rerun()

    with st.expander("➕ Add one worker by hand"):
        _add_form()

    st.markdown("#### Find a worker")
    f1, f2, f3 = st.columns([2, 1, 1])
    q = f1.text_input("Search", placeholder="Name (English or Khmer), worker ID, ID card, phone, NSSF, title", key="hr_q").strip().lower()
    depts = sorted({w["dept"] for w in ws if w["dept"]})
    dep = f2.selectbox("Department", ["All departments"] + depts, key="hr_dep")
    stat = f3.selectbox("Status", ["Working", "Resigned", "All"], key="hr_stat")
    g1, g2, g3 = st.columns(3)
    only_missing = g1.checkbox("Documents missing", key="hr_f_missing")
    only_f = g2.checkbox("Women only", key="hr_f_f")
    only_m = g3.checkbox("Men only", key="hr_f_m")

    def keep(w):
        if stat != "All" and w["status"] != stat: return False
        if dep != "All departments" and w["dept"] != dep: return False
        if only_missing and not w["docs_missing"]: return False
        if only_f and w["sex"] != "F": return False
        if only_m and w["sex"] != "M": return False
        if q and not any(q in str(w[k]).lower() for k in ("id", "name_en", "name_kh", "id_card", "phone", "nssf", "title")): return False
        return True

    show = [w for w in ws if keep(w)]
    if not ws:
        st.info("The worker book is empty. Open the import box above and read your CV sheet.")
        return
    st.caption(f"Showing {len(show)} of {len(ws)}. Click a row to open the profile. Click a column name to sort.")
    if not show:
        st.warning("Nobody matches this search.")
        return
    df = pd.DataFrame([{"ID": w["id"], "Name": w["name_en"], "Khmer name": w["name_kh"], "Dept": w["dept"], "Title": w["title"],
                        "Sex": w["sex"], "Entry date": w["entry_date"], "Status": w["status"],
                        "Docs missing": w["docs_missing"]} for w in show])
    ver = st.session_state.get("hr_list_ver", 0)
    ev = st.dataframe(df, hide_index=True, width="stretch", height=min(60 + 35 * len(df), 620), on_select="rerun",
                      selection_mode="single-row", key=f"hr_list_{ver}")
    rows = list(ev.selection.rows) if ev and ev.selection else []
    if rows:
        st.session_state["hr_open"] = df.iloc[rows[0]]["ID"]
        st.session_state["hr_list_ver"] = ver + 1
        st.rerun()
    st.download_button("⬇️ Download this list (opens in Excel)", _csv(df), file_name="workers.csv", mime="text/csv")


def _add_form():
    with st.form("hr_add", clear_on_submit=False):
        a1, a2, a3 = st.columns(3)
        wid = a1.text_input("Worker ID", placeholder="DMO-10300")
        name_en = a2.text_input("Name (English)")
        name_kh = a3.text_input("Name (Khmer)")
        b1, b2, b3, b4 = st.columns(4)
        dept = b1.text_input("Department")
        title = b2.text_input("Title / position")
        sex = b3.selectbox("Sex", ["", "M", "F"])
        entry = b4.date_input("Entry date", value=None, min_value=MIN_DATE, max_value=MAX_DATE)
        newfp = st.checkbox("🆕 New worker: fingerprint not registered yet", value=True)
        if st.form_submit_button("Add this worker", type="primary"):
            err = H.save_worker({"id": wid, "name_en": name_en, "name_kh": name_kh, "dept": dept, "title": title, "sex": sex,
                                 "entry_date": _iso(entry), "fp_registered": "N" if newfp else "Y"}, new=True)
            if err:
                st.error(err)
            else:
                st.session_state["hr_flash"] = f"{wid.strip().upper()} added. Open the profile to fill the rest."
                st.rerun()


# ================================================================== one worker
def _initials(w):
    parts = (w["name_en"] or w["id"]).replace("-", " ").split()
    return html.escape("".join(p[0] for p in parts[:2]).upper() or "?")


def render_profile(wid):
    w = H.get_worker(wid)
    if not w:
        st.session_state.pop("hr_open", None)
        st.rerun()
        return
    if st.button("⬅️ Back to all workers", key="hr_back"):
        st.session_state.pop("hr_open", None)
        st.rerun()
    flash = st.session_state.pop("hr_flash", "")
    if flash:
        st.success(flash)
    resigned = bool(w["stop_date"] or w["stop_date_raw"])
    docs = H.get_docs(wid)
    missing = [d["doc_type"] for d in docs if not d["given"] and d["doc_type"] in {x["name"] for x in H.doc_types() if x["required"]}]

    with st.container(border=True, key="hr_head"):
        left, right = st.columns([1, 4])
        with left:
            p = H.photo_path(wid)
            if p:
                st.image(p, width=130)
            else:
                st.markdown(f"<div class='hr-avatar'>{_initials(w)}</div>", unsafe_allow_html=True)
        with right:
            st.markdown(f"## {html.escape(w['name_en'] or wid)}")
            if w["name_kh"]:
                st.markdown(f"### {html.escape(w['name_kh'])}")
            st.markdown(f"<div class='hr-sub'><b>{wid}</b> · {html.escape(w['dept'] or '-')} · {html.escape(w['title'] or '-')}" +
                        (f" · started {w['entry_date']}" if w["entry_date"] else "") + "</div>", unsafe_allow_html=True)
            age = H.age_of(w["birthday"])
            pills = ("<span class='hr-pill hr-red'>Resigned</span>" if resigned else "<span class='hr-pill hr-green'>Working</span>")
            pills += (f"<span class='hr-pill hr-amber'>{len(missing)} document(s) missing</span>" if missing else "<span class='hr-pill hr-green'>Documents complete</span>")
            pills += (f"<span class='hr-pill hr-green'>Age {age}</span>" if age != "" else "")
            st.markdown(f"<div style='margin-top:10px'>{pills}</div>", unsafe_allow_html=True)
        pu1, pu2 = st.columns([3, 1])
        up = pu1.file_uploader("Upload a photo", type=["jpg", "jpeg", "png"], key=f"hr_ph_{wid}")
        if up is not None and pu2.button("Save this photo", key=f"hr_phs_{wid}"):
            err = H.save_photo(wid, up.getvalue())
            if err:
                st.error(err)
            else:
                st.session_state["hr_flash"] = "Photo saved."
                st.rerun()

    t1, t2, t5, t3, t6, t4 = st.tabs(["👤 Profile", "📄 Documents", "📑 Contracts", "🗂️ Files", "⚠️ Warnings", "🕘 History"])
    with t1:
        _profile_form(w)
    with t2:
        _docs_tab(wid, docs, missing)
    with t5:
        _contracts_tab(w)
    with t3:
        _files_tab(wid, docs)
    with t6:
        _warnings_tab(wid)
    with t4:
        _history_tab(wid)


def _profile_form(w):
    with st.form(f"hr_prof_{w['id']}"):
        r1 = st.columns(3)
        name_en = r1[0].text_input("Name (English)", w["name_en"])
        name_kh = r1[1].text_input("Name (Khmer)", w["name_kh"])
        dept = r1[2].text_input("Department", w["dept"])
        r2 = st.columns(4)
        title = r2[0].text_input("Title / position", w["title"])
        sex = r2[1].selectbox("Sex", ["", "M", "F"], index=["", "M", "F"].index(w["sex"]) if w["sex"] in ("", "M", "F") else 0)
        entry = r2[2].date_input("Entry date", value=_d(w["entry_date"]), min_value=MIN_DATE, max_value=MAX_DATE)
        birth = r2[3].date_input("Birthday", value=_d(w["birthday"]), min_value=MIN_DATE, max_value=MAX_DATE)
        r3 = st.columns(4)
        id_card = r3[0].text_input("ID card number", w["id_card"])
        phone = r3[1].text_input("Phone", w["phone"])
        nssf = r3[2].text_input("NSSF number", w["nssf"])
        province = r3[3].text_input("Province", w["province"])
        r4 = st.columns(5)
        yn = ["", "Y", "N"]
        married = r4[0].selectbox("Married", yn, index=yn.index(w["married"]) if w["married"] in yn else 0)
        children = r4[1].text_input("Children", w["children"])
        union = r4[2].selectbox("Union member", yn, index=yn.index(w["union_member"]) if w["union_member"] in yn else 0)
        health = r4[3].selectbox("Health check", yn, index=yn.index(w["health_check"]) if w["health_check"] in yn else 0)
        edu = r4[4].text_input("Education", w["education"])
        st.markdown("**Resignation** (fill only when the worker leaves)")
        r5 = st.columns(3)
        stop = r5[0].date_input("Stop date", value=_d(w["stop_date"]), min_value=MIN_DATE, max_value=MAX_DATE)
        stop_kind = r5[1].text_input("Stop kind (from your Excel)", w["stop_kind"])
        stop_amount = r5[2].text_input("Stop amount (from your Excel)", w["stop_amount"])
        if w["stop_date_raw"] and not w["stop_date"]:
            st.warning(f"Your Excel has the text '{w['stop_date_raw']}' as the stop date. It is not a real date. "
                       "Choose the real date above, or leave it empty if the worker did not leave.")
        fpn = st.checkbox("🆕 New worker: fingerprint NOT registered yet (the daily check fills their hours and OT from the paper)", value=(w["fp_registered"] == "N"),
                          help="Tick this for a new worker until the fingerprint machine knows them. Un-tick it as soon as they can scan. While it is ticked, a missing scan is not a warning.")
        remark = st.text_input("Remark from the Excel (section, shift)", w["remark"])
        note = st.text_area("My notes about this worker", w["note"], height=90)
        if st.form_submit_button("💾 Save this profile", type="primary"):
            err = H.save_worker({"id": w["id"], "name_en": name_en, "name_kh": name_kh, "dept": dept, "title": title, "sex": sex,
                                 "entry_date": _iso(entry), "birthday": _iso(birth), "id_card": id_card, "phone": phone, "nssf": nssf,
                                 "province": province, "married": married, "children": children, "union_member": union,
                                 "health_check": health, "education": edu, "stop_date": _iso(stop),
                                 "stop_date_raw": "" if stop else w["stop_date_raw"], "stop_kind": stop_kind,
                                 "stop_amount": stop_amount, "remark": remark, "note": note, "fp_registered": "N" if fpn else "Y"})
            if err:
                st.error(err)
            else:
                st.session_state["hr_flash"] = "Profile saved."
                st.rerun()


def _docs_tab(wid, docs, missing):
    if missing:
        st.warning("Missing: " + ", ".join(missing))
    else:
        st.success("All required documents are given.")
    st.caption("Tick what the worker gave you and type the date. Press Save.")
    df = pd.DataFrame([{"Document": d["doc_type"], "Given": d["given"], "Date given": _d(d["given_on"]), "Note": d["note"]} for d in docs])
    ver = st.session_state.get("hr_docs_ver", 0)
    ed = st.data_editor(df, hide_index=True, width="stretch", disabled=["Document"], key=f"hr_docs_{wid}_{ver}",
                        column_config={"Given": st.column_config.CheckboxColumn(width=70),
                                       "Date given": st.column_config.DateColumn(format="DD MMM YYYY", min_value=MIN_DATE, max_value=MAX_DATE),
                                       "Note": st.column_config.TextColumn(width="large")})
    if st.button("💾 Save the documents", type="primary", key=f"hr_docs_save_{wid}"):
        items, auto = [], 0
        for _, e in ed.iterrows():
            on = e["Date given"]
            on = "" if (on is None or on != on) else str(on)[:10]
            if e["Given"] and not on:
                on, auto = date.today().isoformat(), auto + 1
            items.append({"doc_type": e["Document"], "given": bool(e["Given"]), "given_on": on, "note": str(e["Note"] or "").strip()[:200]})
        H.save_docs(wid, items)
        st.session_state["hr_docs_ver"] = ver + 1
        st.session_state["hr_flash"] = "Documents saved." + (f" {auto} ticked without a date: I used today's date." if auto else "")
        st.rerun()
    n1, n2 = st.columns([3, 1])
    newdoc = n1.text_input("Add a new kind of document for everybody (example: Vaccine card)", key="hr_newdoc")
    if n2.button("Add this document type", key="hr_newdoc_btn") and H.add_doc_type(newdoc):
        st.session_state["hr_flash"] = "Document type added. It is now required for every worker."
        st.rerun()


def _files_tab(wid, docs):
    st.caption("Keep the scans here: CV, ID card, contract, health check. Allowed: PDF, JPG, PNG, Word, Excel, up to 10 MB each.")
    kinds = [d["doc_type"] for d in docs] + ["Other"]
    c1, c2 = st.columns([1, 2])
    kind = c1.selectbox("What is it?", kinds, key=f"hr_fk_{wid}")
    up = c2.file_uploader("Choose a file", key=f"hr_fu_{wid}_{st.session_state.get('hr_fver', 0)}")
    if up is not None and st.button("Save this file", type="primary", key=f"hr_fs_{wid}"):
        err = H.save_file(wid, kind, up.name, up.getvalue())
        if err:
            st.error(err)
        else:
            msg = "File saved."
            cur = {d["doc_type"]: d for d in docs}
            if kind in cur and not cur[kind]["given"]:
                H.save_docs(wid, [{"doc_type": kind, "given": True, "given_on": date.today().isoformat(), "note": cur[kind]["note"]}])
                msg += f" '{kind}' is now ticked as given today. Change the date in Documents if needed."
            st.session_state["hr_fver"] = st.session_state.get("hr_fver", 0) + 1
            st.session_state["hr_flash"] = msg
            st.rerun()
    files = H.list_files(wid)
    if not files:
        st.info("No files yet.")
    for i, f in enumerate(files):
        x1, x2 = st.columns([4, 1])
        x1.markdown(f"**{f['doc']}** · {f['name'].split('__', 2)[-1]} · {f['size_kb']} KB")
        with open(f["path"], "rb") as fh:
            x2.download_button("Open / download", fh.read(), file_name=f["name"].split("__", 2)[-1], key=f"hr_dl_{wid}_{i}")


def _history_tab(wid):
    try:
        hist = db.worker_history(wid)[0]
    except Exception:
        hist = []
    try:
        decs = db.worker_decisions(wid)
    except Exception:
        decs = []
    try:
        lets = [x for x in db.letters_list() if x["worker_id"] == wid]
    except Exception:
        lets = []
    st.markdown("**Letters**")
    if lets:
        st.dataframe(pd.DataFrame([{"Missed-punch date": x["day"], "Type": x.get("kind", ""), "Told on": x["told_on"],
                                    "Letter back on": x["returned_on"], "I told the result on": x.get("replied_on", ""), "Note": x["note"]} for x in lets]),
                     hide_index=True, width="stretch")
    else:
        st.caption("No letters.")
    st.markdown("**What you decided on attendance days**")
    if decs:
        st.dataframe(pd.DataFrame([{"Day": x["day"], "Hours": x["hours"], "OT": x["ot"], "Code": x["code"], "Reason": x["reason"]} for x in decs]),
                     hide_index=True, width="stretch")
    else:
        st.caption("Nothing saved yet.")
    st.markdown("**Attendance remarks**")
    if hist:
        st.dataframe(pd.DataFrame([{k: v for k, v in x.items() if k in ("day", "status", "paper_ot", "finger_ot", "in_time", "out_time", "remark")} for x in hist]),
                     hide_index=True, width="stretch")
    else:
        st.caption("No attendance remarks saved yet.")


# ================================================================== contracts (every 3 months)
def _tx(v):
    return "" if (v is None or v != v) else str(v)[:10]


def _contracts_tab(w):
    wid = w["id"]
    cs = H.contracts_of(wid)
    st.caption("A new 3-month contract again and again (whole months, after 2 months of probation). Type the dates: made, signed, given back, and the day you paid the 5%. Press Save.")
    if not w["entry_date"]:
        st.warning("This worker has no entry date. Type it in the Profile tab, then I can suggest the contract periods.")
    if cs:
        df = pd.DataFrame([{"No.": c["no"], "Start": _d(c["start"]), "End": _d(c["end"]), "Made on": _d(c["made_on"]), "Signed on": _d(c["signed_on"]),
                            "Given back on": _d(c["returned_on"]), "5% paid on": _d(c.get("settled_on", "")), "Note": c["note"], "Delete": False} for c in cs])
        ver = st.session_state.get("hr_c_ver", 0)
        ed = st.data_editor(df, hide_index=True, width="stretch", disabled=["No."], key=f"hr_ct_{wid}_{ver}",
                            column_config={k: st.column_config.DateColumn(format="DD MMM YYYY", min_value=MIN_DATE, max_value=MAX_DATE)
                                           for k in ("Start", "End", "Made on", "Signed on", "Given back on", "5% paid on")} |
                                          {"Delete": st.column_config.CheckboxColumn(width=60)})
        if st.button("💾 Save the contracts", type="primary", key=f"hr_ct_save_{wid}"):
            errs = []
            for c, (_, e) in zip(cs, ed.iterrows()):
                if e["Delete"]:
                    H.delete_contract(c["id"])
                    continue
                err = H.save_contract({"worker_id": wid, "no": c["no"], "start": _tx(e["Start"]), "end": _tx(e["End"]), "made_on": _tx(e["Made on"]),
                                       "signed_on": _tx(e["Signed on"]), "returned_on": _tx(e["Given back on"]), "settled_on": _tx(e["5% paid on"]), "note": str(e["Note"] or "").strip()})
                if err:
                    errs.append(err)
            if errs:
                st.error("Not saved:\n\n" + "\n\n".join(errs))
            else:
                st.session_state["hr_c_ver"] = ver + 1
                st.session_state["hr_flash"] = "Contracts saved."
                st.rerun()
    else:
        st.info("No contract saved for this worker yet.")
    if w["entry_date"]:
        entry = _d(w["entry_date"])
        nxt = (max([c["no"] for c in cs]) + 1) if cs else max(1, H.period_no_on(entry, date.today()))
        s, e = H.period(entry, nxt)
        if st.button(f"➕ Add contract {nxt} ({s:%d %b %Y} to {e:%d %b %Y})", key=f"hr_ct_add_{wid}"):
            err = H.save_contract({"worker_id": wid, "no": nxt, "start": s.isoformat(), "end": e.isoformat()})
            st.session_state["hr_flash"] = err or f"Contract {nxt} added. Type the dates when you make it."
            st.rerun()
    if cs:
        st.markdown("**Scan of a signed contract**")
        k1, k2 = st.columns([1, 2])
        no = k1.selectbox("Which contract", [c["no"] for c in cs], key=f"hr_cs_no_{wid}")
        up = k2.file_uploader("Choose the scan", key=f"hr_cs_up_{wid}_{st.session_state.get('hr_fver', 0)}")
        if up is not None and st.button("Save this scan", key=f"hr_cs_save_{wid}"):
            err = H.save_file(wid, f"Contract {no}", up.name, up.getvalue())
            st.session_state["hr_fver"] = st.session_state.get("hr_fver", 0) + 1
            st.session_state["hr_flash"] = err or f"Scan saved for contract {no}. You see it in the Files tab."
            st.rerun()


def render_contracts():
    _theme()
    st.markdown(f"<div class='hr-sub'>{date.today():%A, %d %B %Y}</div>", unsafe_allow_html=True)
    st.title("Contracts")
    st.markdown("<div class='hr-sub'>Probation = the month the worker entered + the next month. Then contract 1 starts on the 1st of the third month, "
                "and a new 3-month contract follows again and again (entered 27 Jun: probation Jun-Jul, contract 1 = 1 Aug to 31 Oct). "
                "When a contract ends you pay the 5% (\u7ed3\u7b97). <b>Nothing is saved until you type a date and press Save.</b></div>",
                unsafe_allow_html=True)
    month_todo_block()
    flash = st.session_state.pop("hr_flash", "")
    if flash:
        st.success(flash)
    no_entry = [w for w in H.list_workers() if w["status"] == "Working" and not w["entry_date"]]
    if no_entry:
        st.warning(f"{len(no_entry)} working worker(s) have no entry date, so I cannot suggest their contracts.")

    show = st.radio("Show", ["Starts in 14 days", "Starts in 30 days", "Starts in 60 days", "Only contracts I already made (not finished)"],
                    index=1, horizontal=True, key="hr_ct_show")
    days = {"Starts in 14 days": 14, "Starts in 30 days": 30, "Starts in 60 days": 60}.get(show, 60)
    rows = H.contract_rows(within_days=days)
    if show.startswith("Only"):
        rows = [r for r in rows if r["state"] == "open"]
    allr = H.contract_rows(within_days=30)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Start in 30 days, not made", sum(r["state"] == "due" for r in allr))
    k2.metric("Running without a made contract", sum(r["state"] == "late" for r in allr))
    k3.metric("Made, not signed yet", sum(r["state"] == "open" and not r["signed_on"] for r in allr))
    k4.metric("Signed, not given back", sum(bool(r["state"] == "open" and r["signed_on"] and not r["returned_on"]) for r in allr))

    f1, f2 = st.columns([1, 2])
    depts = sorted({r["dept"] for r in rows if r["dept"]})
    dep = f1.selectbox("Department", ["All departments"] + depts, key="hr_ct_dep")
    q = f2.text_input("Search", placeholder="Name or worker ID", key="hr_ct_q").strip().lower()
    rows = [r for r in rows if (dep == "All departments" or r["dept"] == dep) and (not q or q in r["worker_id"].lower() or q in (r["name"] or "").lower() or q in (r["name_kh"] or ""))]
    if not rows:
        st.success("Nothing to do in this list.")
        settle_block()
        return

    def status(r):
        if r["state"] == "late":
            return "Running, no contract made"
        if r["state"] == "open":
            return "Signed, not given back" if r["signed_on"] else "Made, not signed"
        d = r["days_to_start"]
        return "Starts today" if d == 0 else (f"Starts in {d} days" if d > 0 else f"Started {-d} days ago")
    df = pd.DataFrame([{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Contract": f"Contract {r['no']}", "Status": status(r),
                        "Start": r["start"], "End": r["end"], "Made on": _d(r["made_on"]), "Signed on": _d(r["signed_on"]),
                        "Given back on": _d(r["returned_on"]), "Note": r["note"]} for r in rows])
    ver = st.session_state.get("hr_ctl_ver", 0)
    dc = st.column_config.DateColumn(format="DD MMM YYYY", min_value=MIN_DATE, max_value=MAX_DATE)
    ed = st.data_editor(df, hide_index=True, width="stretch", height=min(60 + 36 * len(df), 640), key=f"hr_ctl_{ver}_{len(df)}_{hash(tuple(df['Worker'])) % 99999}",
                        disabled=["Worker", "Name", "Dept", "Contract", "Status"],
                        column_config={"Start": dc, "End": dc, "Made on": dc, "Signed on": dc, "Given back on": dc,
                                       "Note": st.column_config.TextColumn(width="medium")})
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        if st.button("💾 Save the dates", type="primary", key="hr_ctl_save"):
            errs, n = [], 0
            for r, (_, e) in zip(rows, ed.iterrows()):
                made, signed, back = _tx(e["Made on"]), _tx(e["Signed on"]), _tx(e["Given back on"])
                note = str(e["Note"] or "").strip()
                if not (made or signed or back or note or r["stored"]):
                    continue
                err = H.save_contract({"worker_id": r["worker_id"], "no": r["no"], "start": _tx(e["Start"]), "end": _tx(e["End"]),
                                       "made_on": made, "signed_on": signed, "returned_on": back, "note": note})
                if err:
                    errs.append(err)
                else:
                    n += 1
            if errs:
                st.error("Some lines were not saved:\n\n" + "\n\n".join(errs))
            if n:
                st.session_state["hr_ctl_ver"] = ver + 1
                st.session_state["hr_flash"] = f"Saved {n} contract line(s)." + (f" {len(errs)} line(s) had a problem." if errs else "")
                st.rerun()
    with c2:
        todo = [r for r in rows if not r["made_on"]]
        if todo and st.button(f"Mark 'Made' = today for the {len(todo)} shown without a made date", key="hr_ctl_today"):
            for r in todo:
                H.save_contract({"worker_id": r["worker_id"], "no": r["no"], "start": r["start"].isoformat(), "end": r["end"].isoformat(),
                                 "made_on": date.today().isoformat(), "note": r["note"]})
            st.session_state["hr_ctl_ver"] = ver + 1
            st.session_state["hr_flash"] = f"{len(todo)} contract(s) marked as made today."
            st.rerun()
    with c3:
        out = df.copy()
        for c in ("Start", "End", "Made on", "Signed on", "Given back on"):
            out[c] = out[c].map(lambda v: "" if v is None else f"{v:%Y-%m-%d}")
        st.download_button("⬇️ Download", _csv(out), file_name="contracts.csv", mime="text/csv")
    settle_block()


def _xlsx_safe(v):
    return "'" + v if isinstance(v, str) and v[:1] in ("=", "+", "-", "@", "\t", "\r") else v


def settle_block():
    """Contracts that ended or end soon: mark the day you paid the 5% (\u7ed3\u7b97)."""
    rows = H.settlement_rows(within_days=30)
    with st.expander(f"💵 5% end-of-contract money (\u7ed3\u7b97): {len(rows)} to pay or check", expanded=bool(rows)):
        if not rows:
            st.success("Nothing to pay in the next 30 days.")
            return
        st.caption("Type the day you paid the 5% for that contract, then Save. Red = the contract already ended.")
        df = pd.DataFrame([{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Contract": f"Contract {r['no']}", "Ends": r["end"],
                            "Days left": r["days_left"], "5% paid on": None} for r in rows])
        ver = st.session_state.get("hr_st_ver", 0)
        ed = st.data_editor(df, hide_index=True, width="stretch", disabled=["Worker", "Name", "Dept", "Contract", "Ends", "Days left"], key=f"hr_st_{ver}_{len(df)}",
                            column_config={"Ends": st.column_config.DateColumn(format="DD MMM YYYY"),
                                           "5% paid on": st.column_config.DateColumn(format="DD MMM YYYY", min_value=MIN_DATE, max_value=MAX_DATE)})
        if st.button("💾 Save the 5% dates", type="primary", key="hr_st_save"):
            n, errs = 0, []
            for r, (_, e) in zip(rows, ed.iterrows()):
                paid = _tx(e["5% paid on"])
                if not paid:
                    continue
                mine = {c["no"]: c for c in H.contracts_of(r["worker_id"])}.get(r["no"], {})
                err = H.save_contract({"worker_id": r["worker_id"], "no": r["no"], "start": r["start"].isoformat(), "end": r["end"].isoformat(), "settled_on": paid,
                                       "made_on": mine.get("made_on", ""), "signed_on": mine.get("signed_on", ""), "returned_on": mine.get("returned_on", ""), "note": mine.get("note", "")})
                if err:
                    errs.append(err)
                else:
                    n += 1
            if errs:
                st.error("Not saved:\n\n" + "\n\n".join(errs))
            if n:
                st.session_state["hr_st_ver"] = ver + 1
                st.session_state["hr_flash"] = f"Saved {n} date(s)."
                st.rerun()


TEMPLATE_PATH = os.path.join("memory", "contract_template.docx")


def month_todo_block():
    """The list of what you must do in one month, as an Excel you can download, and the contract papers filled for the workers you choose."""
    today = date.today()
    with st.expander("📋 What I must do in a month (list to download + contract papers)", expanded=True):
        c1, c2 = st.columns(2)
        yr = c1.number_input("Year", 2020, 2100, today.year, key="hr_mt_y")
        mo = c2.selectbox("Month", list(range(1, 13)), index=today.month - 1, format_func=lambda m: date(2000, m, 1).strftime("%B"), key="hr_mt_m")
        make, settle, prob = H.month_todo(int(yr), int(mo))
        k = st.columns(3)
        k[0].metric("Contracts to make", len(make))
        k[1].metric("Pay the 5% (contract ends)", len(settle))
        k[2].metric("Probation ends", len(prob))
        t1 = pd.DataFrame([{"Worker": r["worker_id"], "Name": r["name"], "Khmer name": r["name_kh"], "Dept": r["dept"], "Position": r["title"], "Entry": r["entry"].isoformat(),
                            "Contract": r["no"], "Starts": r["start"].isoformat(), "Ends": r["end"].isoformat(), "Made on": r["made_on"],
                            "Signed on": r["signed_on"], "Given back on": r["returned_on"]} for r in make])
        t2 = pd.DataFrame([{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Contract": r["no"], "Ends": r["end"].isoformat(), "5% paid on": r["settled_on"]} for r in settle])
        t3 = pd.DataFrame([{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Entry": r["entry"].isoformat(), "Probation ends": r["probation_end"].isoformat(),
                            "Contract 1 starts": (r["probation_end"] + timedelta(days=1)).isoformat()} for r in prob])
        a, b, c = st.tabs([f"Make ({len(t1)})", f"Pay the 5% ({len(t2)})", f"Probation ends ({len(t3)})"])
        with a:
            st.dataframe(t1, hide_index=True, width="stretch", height=min(60 + 35 * max(len(t1), 1), 420))
        with b:
            st.dataframe(t2, hide_index=True, width="stretch", height=min(60 + 35 * max(len(t2), 1), 420))
        with c:
            st.dataframe(t3, hide_index=True, width="stretch", height=min(60 + 35 * max(len(t3), 1), 420))
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as xw:
            for name, df in (("Make contracts", t1), ("Pay the 5%", t2), ("Probation ends", t3)):
                df.map(_xlsx_safe).to_excel(xw, sheet_name=name, index=False)
        st.download_button("⬇️ Download this month's list (Excel)", bio.getvalue(), file_name=f"contracts_to_do_{int(yr)}-{int(mo):02d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="hr_mt_dl")
        st.markdown("**📄 Fill the contract paper**")
        have = os.path.exists(TEMPLATE_PATH)
        up = st.file_uploader("Your contract paper (Word .docx). Upload it once; I only read it and make new files.", type=["docx"], key="hr_tpl_up")
        if up is not None and st.button("Save this paper as my template", key="hr_tpl_save"):
            if len(up.getvalue()) > 3 * 1024 * 1024:
                st.error("The file is bigger than 3 MB.")
            else:
                os.makedirs("memory", exist_ok=True)
                with open(TEMPLATE_PATH, "wb") as f:
                    f.write(up.getvalue())
                st.session_state["hr_flash"] = "Template saved."
                st.rerun()
        if not have:
            st.info("Upload your contract paper above first (the one with a worker's data in it is fine; I replace the data).")
            return
        st.caption("The paper in use: memory/contract_template.docx. Pay: salary and allowance are the same for everybody you choose below. Birth place and address are filled from 'Province' of the worker book.")
        s1, s2 = st.columns(2)
        sal = s1.number_input("Base salary (USD)", 0, 5000, H.get_setting("contract_salary"), key="hr_mt_sal")
        alw = s2.number_input("Position allowance (USD)", 0, 5000, H.get_setting("contract_allowance"), key="hr_mt_alw")
        opts = {f"{r['worker_id']} · {r['name']} · contract {r['no']}": r for r in make}
        pick = st.multiselect("Choose the workers (empty = nobody)", list(opts), default=list(opts)[:0], key="hr_mt_pick")
        if st.button("📄 Make the papers", type="primary", key="hr_mt_make", disabled=not pick):
            try:
                import contract_docx as CD
            except ImportError:
                st.error("The program python-docx is not installed. Close the website, open a terminal in D:\\attendance and type:  py -m pip install python-docx   then start the website again.")
                return
            H.set_setting("contract_salary", sal); H.set_setting("contract_allowance", alw)
            tpl = open(TEMPLATE_PATH, "rb").read()
            files, errs = {}, []
            ws = {w["id"]: w for w in H.list_workers()}
            for key in pick:
                r = opts[key]
                try:
                    files[f"{r['worker_id']}_contract{r['no']}.docx"] = CD.fill(tpl, ws[r["worker_id"]], {"start": r["start"], "end": r["end"]}, int(sal), int(alw))
                except Exception as e:
                    errs.append(f"{r['worker_id']}: {e}")
            if errs:
                st.error("Some papers were not made:\n\n" + "\n\n".join(errs))
            if files:
                zb = io.BytesIO()
                with zipfile.ZipFile(zb, "w", zipfile.ZIP_DEFLATED) as z:
                    for n_, b_ in files.items():
                        z.writestr(n_, b_)
                st.session_state["hr_mt_zip"] = (zb.getvalue(), f"contracts_{int(yr)}-{int(mo):02d}.zip", len(files))
        z = st.session_state.get("hr_mt_zip")
        if z:
            st.success(f"{z[2]} contract paper(s) are ready. Open each one and CHECK it before you print (missing data stays as it was in your paper).")
            st.download_button("⬇️ Download the papers (zip)", z[0], file_name=z[1], mime="application/zip", key="hr_mt_zipdl")


# ================================================================== shared helpers for the new pages
def _worker_options(working_only=True):
    ws = [w for w in H.list_workers() if (w["status"] == "Working" or not working_only)]
    return [f"{w['id']} - {w['name_en'] or w['name_kh'] or ''}" for w in ws]


def _wid_of(option):
    return str(option).split(" - ")[0].strip()


def _flash():
    f = st.session_state.pop("hr_flash", "")
    if f:
        st.success(f)
    w = st.session_state.pop("hr_flash_warn", "")
    if w:
        st.warning(w)


def _head(title, sub=""):
    _theme()
    st.markdown(f"<div class='hr-sub'>{date.today():%A, %d %B %Y}</div>", unsafe_allow_html=True)
    st.title(title)
    if sub:
        st.markdown(f"<div class='hr-sub'>{sub}</div>", unsafe_allow_html=True)
    _flash()


_DC = st.column_config.DateColumn(format="DD MMM YYYY", min_value=MIN_DATE, max_value=MAX_DATE)


def _warnings_tab(wid):
    st.caption("Warning letters given to this worker. Add one when you give a letter.")
    c1, c2, c3 = st.columns([1, 1, 2])
    on = c1.date_input("Date of the letter", value=date.today(), min_value=MIN_DATE, max_value=MAX_DATE, key=f"hr_wn_d_{wid}")
    kind = c2.selectbox("Kind", H.WARNING_KINDS, key=f"hr_wn_k_{wid}")
    why = c3.text_input("Reason", key=f"hr_wn_r_{wid}")
    if st.button("Add this warning letter", key=f"hr_wn_add_{wid}"):
        err = H.add_warning(wid, _iso(on), kind, why)
        st.session_state["hr_flash"] = err or "Warning letter saved."
        st.rerun()
    ws = H.list_warnings(wid)
    if not ws:
        st.info("No warning letters for this worker.")
    for x in ws:
        a, b = st.columns([6, 1])
        a.markdown(f"**{x['on_date']}** · {x['kind']} · {html.escape(x['reason'] or '')}", unsafe_allow_html=True)
        if b.button("Delete", key=f"hr_wn_del_{x['id']}"):
            H.delete_warning(x["id"])
            st.rerun()


# ================================================================== TODAY
def _alert_block(a):
    icon = {"red": "🔴", "amber": "🟡", "green": "🟢"}[a["level"]]
    with st.expander(f"{icon}  {a['title']}  ·  {len(a['rows'])}", expanded=(a["level"] == "red")):
        st.dataframe(pd.DataFrame(a["rows"]), hide_index=True, width="stretch", height=min(60 + 35 * len(a["rows"]), 360))
        if a["where"]:
            st.caption(f"Go to: HR > {a['where']}" if a["where"] not in ("Daily work", "Letters") else f"Go to: Attendance > {a['where']}")


def render_today():
    _head("Today", "What needs you first. Red = late or urgent. Yellow = do soon. Green = good news.")
    ws = H.list_workers()
    working = [w for w in ws if w["status"] == "Working"]
    today = date.today()
    left_month = sum(1 for w in ws if H._d(w["stop_date"]) and H._d(w["stop_date"]).year == today.year and H._d(w["stop_date"]).month == today.month)
    toreq = sum(q["left"] for q in H.list_requests(include_closed=False))
    if not ws:
        st.info("The worker book is empty. Go to HR > Workers and import your CV sheet first.")
    k = st.columns(5)
    k[0].metric("Working now", len(working))
    k[1].metric("Women", sum(w["sex"] == "F" for w in working))
    k[2].metric("Men", sum(w["sex"] == "M" for w in working))
    k[3].metric("Still to recruit", toreq)
    k[4].metric("Resigned this month", left_month)

    al = H.alerts()
    nr, na = sum(a["level"] == "red" for a in al), sum(a["level"] == "amber" for a in al)
    st.markdown("#### Needs you")
    if not (nr or na):
        st.success("Nothing urgent today.")
    else:
        st.markdown(f"<span class='hr-pill hr-red'>{nr} red</span><span class='hr-pill hr-amber'>{na} yellow</span>", unsafe_allow_html=True)
    for a in al:
        _alert_block(a)

    st.markdown("#### Departments: planned, real, to recruit")
    st.caption("Type the planned number of workers for each department and press Save. 'Asked' is what open recruitment requests still need.")
    rows = H.dept_table(ws)
    if rows:
        df = pd.DataFrame([{"Department": r["dept"], "Planned": r["planned"], "Real": r["real"], "Gap": r["gap"], "Asked (open requests)": r["asked"]} for r in rows])
        ver = st.session_state.get("hr_plan_ver", 0)
        ed = st.data_editor(df, hide_index=True, width="stretch", key=f"hr_plan_{ver}",
                            disabled=["Department", "Real", "Gap", "Asked (open requests)"],
                            column_config={"Planned": st.column_config.NumberColumn(min_value=0, max_value=5000, step=1)})
        if st.button("💾 Save planned numbers", type="primary", key="hr_plan_save"):
            H.save_plan({e["Department"]: e["Planned"] for _, e in ed.iterrows() if e["Planned"] == e["Planned"]})
            st.session_state["hr_plan_ver"] = ver + 1
            st.session_state["hr_flash"] = "Planned numbers saved."
            st.rerun()
        chart = df.set_index("Department")[["Planned", "Real"]]
        if chart["Planned"].sum() or chart["Real"].sum():
            st.bar_chart(chart, height=260, color=["#E8B04A", "#1F6F5F"])
    n1, n2 = st.columns([3, 1])
    nd = n1.text_input("Add a department that has no worker yet", key="hr_newdept")
    if n2.button("Add department", key="hr_newdept_btn") and nd.strip():
        H.save_plan({nd.strip(): 0})
        st.rerun()

    dashboard_blocks()

    with st.expander("⚙️ Limits I used (change them if your factory uses other numbers)"):
        st.caption("These are my starting numbers, not law. Please check them with your own rule.")
        c = st.columns(5)
        vals = {
            "unpaid_red_days": c[0].number_input("Resigned unpaid: red after (days)", 0, 365, H.get_setting("unpaid_red_days"), key="hr_s1"),
            "track_resign_days": c[1].number_input("Track resigned workers for (days)", 1, 3650, H.get_setting("track_resign_days"), key="hr_s2"),
            "absent_days": c[2].number_input("Absent in a row warning (days)", 1, 30, H.get_setting("absent_days"), key="hr_s3"),
            "fdc_months": c[3].number_input("Fixed-contract limit (months)", 1, 120, H.get_setting("fdc_months"), key="hr_s4"),
            "fdc_warn_days": c[4].number_input("Warn before the limit (days)", 0, 365, H.get_setting("fdc_warn_days"), key="hr_s5")}
        if st.button("Save the limits", key="hr_s_save"):
            for kk, vv in vals.items():
                H.set_setting(kk, vv)
            st.session_state["hr_flash"] = "Limits saved."
            st.rerun()


def dashboard_blocks():
    st.markdown("#### Day and night shift")
    rows = H.shift_counts()
    if not rows:
        st.info("No day is saved yet. When you check a day in 'Daily work' (Step 3) it is saved, and the numbers appear here.")
    else:
        st.caption("Counted from the days you checked: a worker counts when he has a fingerprint scan that work day. "
                   "The 16:00-00:00 evening shift counts as night. Office staff are counted apart.")
        days = [r["day"] for r in rows]
        pick = st.selectbox("Day", days, key="hr_sh_day", format_func=lambda d_: f"{_d(d_):%a %d %b %Y}")
        cur = next(r for r in rows if r["day"] == pick)
        k = st.columns(4)
        k[0].metric("☀️ Day shift", cur["day_n"])
        k[1].metric("🌙 Night shift", cur["night_n"])
        k[2].metric("🏢 Office", cur["office_n"])
        k[3].metric("Total who came", cur["total"])
        with st.expander("Who came on this day, by department"):
            st.dataframe(pd.DataFrame([{"Department": dp, "Day": v["day_n"], "Night": v["night_n"], "Office": v["office_n"]} for dp, v in sorted(cur["by_dept"].items())]),
                         hide_index=True, width="stretch")
        n = len(rows)
        g = st.columns(4)
        g[0].metric("Days saved", n)
        g[1].metric("Average day shift", f"{sum(r['day_n'] for r in rows) / n:.0f}")
        g[2].metric("Average night shift", f"{sum(r['night_n'] for r in rows) / n:.0f}")
        g[3].metric("Most on night", max(r["night_n"] for r in rows))
        months = sorted({r["day"][:7] for r in rows}, reverse=True)
        sel = st.selectbox("Show", ["All saved days"] + months, key="hr_sh_month", format_func=lambda m_: m_ if m_.startswith("All") else f"{date(int(m_[:4]), int(m_[5:7]), 1):%B %Y}")
        sub = [r for r in rows if sel.startswith("All") or r["day"].startswith(sel)]
        df = pd.DataFrame([{"Date": r["day"], "Day": r["day_n"], "Night": r["night_n"], "Office": r["office_n"], "Total": r["total"]} for r in reversed(sub)])
        st.bar_chart(df.set_index("Date")[["Day", "Night", "Office"]], height=260, color=["#E8B04A", "#26302C", "#1F6F5F"])
        st.dataframe(df.iloc[::-1], hide_index=True, width="stretch", height=min(60 + 35 * len(df), 380))
        st.download_button("⬇️ Download (opens in Excel)", _csv(df), file_name=f"day_night_workers_{sel if not sel.startswith('All') else 'all'}.csv", mime="text/csv", key="hr_sh_dl")

    st.markdown("#### Hired and resigned, month by month")
    mv, hired, left = H.monthly_moves()
    if not mv:
        st.info("No entry or stop dates in the worker book yet.")
        return
    st.caption("Hired = the entry date in the worker book. Resigned = the stop date. 'Working at the end' counts everybody who had started and not stopped on the last day of the month.")
    span = st.radio("Months", ["Last 12", "Last 24", "All"], horizontal=True, key="hr_mv_span")
    show = mv[-12:] if span == "Last 12" else (mv[-24:] if span == "Last 24" else mv)
    df = pd.DataFrame([{"Month": r["month"], "Hired": r["hired"], "Resigned": r["left"], "Change": r["net"], "Working at the end": r["working_end"]} for r in show])
    g = st.columns(3)
    g[0].metric("Hired (shown months)", int(df["Hired"].sum()))
    g[1].metric("Resigned (shown months)", int(df["Resigned"].sum()))
    g[2].metric("Change", f"{int(df['Change'].sum()):+d}")
    st.bar_chart(df.set_index("Month")[["Hired", "Resigned"]], height=260, color=["#1F6F5F", "#C0553F"])
    st.dataframe(df.iloc[::-1], hide_index=True, width="stretch", height=min(60 + 35 * len(df), 420))
    st.download_button("⬇️ Download (opens in Excel)", _csv(df), file_name="hired_resigned_by_month.csv", mime="text/csv", key="hr_mv_dl")
    mk = st.selectbox("See the names of one month", [r["month"] for r in reversed(show)], key="hr_mv_pick", format_func=lambda m_: f"{date(int(m_[:4]), int(m_[5:7]), 1):%B %Y}")
    a, b = st.columns(2)
    with a:
        st.markdown("**Hired**")
        st.dataframe(pd.DataFrame([{"ID": w["id"], "Name": w["name_en"], "Dept": w["dept"], "Entry": w["entry_date"]} for w in hired.get(mk, [])]), hide_index=True, width="stretch")
    with b:
        st.markdown("**Resigned**")
        st.dataframe(pd.DataFrame([{"ID": w["id"], "Name": w["name_en"], "Dept": w["dept"], "Left": w["stop_date"], "Why": w["stop_kind"]} for w in left.get(mk, [])]), hide_index=True, width="stretch")


# ================================================================== RECRUITMENT
def render_recruitment():
    _head("Recruitment", "Who asked, how many they need, who you found, and who started. When a person is 'Hired', add them to the worker book.")
    reqs = H.list_requests()
    cands = H.list_candidates()
    openq = [q for q in reqs if q["open"]]
    k = st.columns(4)
    k[0].metric("Open requests", len(openq))
    k[1].metric("Still to recruit", sum(q["left"] for q in openq))
    k[2].metric("Candidates waiting", sum(c["status"] in ("Found", "Interview") for c in cands))
    k[3].metric("Hired, not in worker book", sum(c["status"] == "Hired" and not c["worker_id"] for c in cands))

    with st.expander("➕ New request (someone asked me to recruit)", expanded=not reqs):
        depts = sorted({w["dept"] for w in H.list_workers() if w["dept"]} | set(H.get_plan()))
        with st.form("hr_new_req", clear_on_submit=True):
            r1 = st.columns(3)
            dsel = r1[0].selectbox("Department", depts + ["(new department)"] if depts else ["(new department)"])
            dnew = r1[1].text_input("New department name (only if you chose new)")
            pos = r1[2].text_input("Position (example: Machine operator)")
            r2 = st.columns(4)
            need = r2[0].number_input("How many people", 1, 5000, 1)
            by = r2[1].text_input("Who asked")
            asked = r2[2].date_input("Date asked", value=date.today(), min_value=MIN_DATE, max_value=MAX_DATE)
            dl = r2[3].date_input("Needed by", value=None, min_value=MIN_DATE, max_value=MAX_DATE)
            note = st.text_input("Note")
            if st.form_submit_button("Save this request", type="primary"):
                err = H.save_request({"dept": (dnew if dsel == "(new department)" else dsel), "position": pos, "need": need, "asked_by": by,
                                      "asked_on": _iso(asked), "deadline": _iso(dl), "note": note})
                if err:
                    st.error(err)
                else:
                    st.session_state["hr_flash"] = "Request saved."
                    st.rerun()

    st.markdown("#### Requests")
    if not reqs:
        st.info("No request yet.")
    else:
        df = pd.DataFrame([{"Position": q["position"], "Dept": q["dept"], "Need": q["need"], "Hired": q["hired"], "Still need": q["left"],
                            "Asked by": q["asked_by"], "Asked on": H._d(q["asked_on"]), "Needed by": H._d(q["deadline"]),
                            "Note": q["note"], "Closed": bool(q["closed"]), "Delete": False} for q in reqs])
        ver = st.session_state.get("hr_rq_ver", 0)
        ed = st.data_editor(df, hide_index=True, width="stretch", key=f"hr_rq_{ver}_{len(df)}",
                            disabled=["Hired", "Still need"],
                            column_config={"Asked on": _DC, "Needed by": _DC, "Need": st.column_config.NumberColumn(min_value=1, max_value=5000, step=1),
                                           "Closed": st.column_config.CheckboxColumn(width=70), "Delete": st.column_config.CheckboxColumn(width=70)})
        if st.button("💾 Save the requests", type="primary", key="hr_rq_save"):
            errs = []
            for q, (_, e) in zip(reqs, ed.iterrows()):
                if e["Delete"]:
                    H.delete_request(q["id"])
                    continue
                err = H.save_request({"id": q["id"], "dept": e["Dept"], "position": e["Position"], "need": e["Need"], "asked_by": e["Asked by"] or "",
                                      "asked_on": _tx(e["Asked on"]), "deadline": _tx(e["Needed by"]), "note": e["Note"] or "", "closed": e["Closed"]})
                if err:
                    errs.append(f"{q['position']}: {err}")
            if errs:
                st.error("Not saved:\n\n" + "\n\n".join(errs))
            else:
                st.session_state["hr_rq_ver"] = ver + 1
                st.session_state["hr_flash"] = "Requests saved."
                st.rerun()

    st.markdown("#### Candidates")
    qopt = {f"{q['position']} - {q['dept']} (need {q['need']}, hired {q['hired']})": q["id"] for q in reqs if not q["closed"]}
    if not qopt:
        st.info("Make a request first, then add the people you find for it.")
    else:
        with st.form("hr_new_cand", clear_on_submit=True):
            a = st.columns(3)
            which = a[0].selectbox("For which request", list(qopt))
            en = a[1].text_input("Name (English)")
            kh = a[2].text_input("Name (Khmer)")
            b = st.columns(4)
            ph = b[0].text_input("Phone")
            idc = b[1].text_input("ID card number (to check if this person worked here before)")
            sx = b[2].selectbox("Sex", ["", "M", "F"])
            fo = b[3].date_input("Date found", value=date.today(), min_value=MIN_DATE, max_value=MAX_DATE)
            nt = st.text_input("Note")
            if st.form_submit_button("Add this person", type="primary"):
                err, warns = H.add_candidate({"request_id": qopt[which], "name_en": en, "name_kh": kh, "phone": ph, "id_card": idc, "sex": sx, "found_on": _iso(fo), "note": nt})
                if err:
                    st.error(err)
                else:
                    st.session_state["hr_flash"] = "Person added."
                    if warns:
                        st.session_state["hr_flash_warn"] = "⚠️ Same ID card number is already known:\n\n" + "\n\n".join("- " + x for x in warns)
                    st.rerun()

    if cands:
        f1, f2 = st.columns([2, 1])
        fsel = f1.selectbox("Show candidates of", ["All requests"] + list(qopt), key="hr_cf")
        fst = f2.selectbox("Status", ["All", "Still waiting (Found, Interview)"] + H.CANDIDATE_STATUS, key="hr_cs")
        show = [c for c in cands if (fsel == "All requests" or c["request_id"] == qopt.get(fsel)) and
                (fst == "All" or (fst.startswith("Still") and c["status"] in ("Found", "Interview")) or c["status"] == fst)]
        if show:
            df = pd.DataFrame([{"Name": c["name_en"] or c["name_kh"], "For": f"{c['position'] or '-'} ({c['dept'] or '-'})", "Found on": H._d(c["found_on"]),
                                "Phone": c["phone"], "Status": c["status"], "Interview on": H._d(c["interview_on"]), "Starts on": H._d(c["start_on"]),
                                "Worker ID": c["worker_id"], "Note": c["note"], "Delete": False} for c in show])
            ver = st.session_state.get("hr_cd_ver", 0)
            ed = st.data_editor(df, hide_index=True, width="stretch", key=f"hr_cd_{ver}_{len(df)}_{hash(tuple(df['Name'])) % 9999}",
                                disabled=["Name", "For", "Found on", "Worker ID"],
                                column_config={"Status": st.column_config.SelectboxColumn(options=H.CANDIDATE_STATUS, required=True),
                                               "Interview on": _DC, "Starts on": _DC, "Delete": st.column_config.CheckboxColumn(width=70)})
            if st.button("💾 Save the candidates", type="primary", key="hr_cd_save"):
                items = []
                for c, (_, e) in zip(show, ed.iterrows()):
                    if e["Delete"]:
                        if not c["worker_id"]:
                            H.delete_candidate(c["id"])
                        continue
                    if c["worker_id"]:
                        continue
                    items.append({"id": c["id"], "status": e["Status"], "interview_on": _tx(e["Interview on"]), "start_on": _tx(e["Starts on"]),
                                  "phone": e["Phone"] or "", "note": e["Note"] or ""})
                err = H.update_candidates(items)
                if err:
                    st.error(err)
                else:
                    st.session_state["hr_cd_ver"] = ver + 1
                    st.session_state["hr_flash"] = "Candidates saved."
                    st.rerun()
        else:
            st.info("Nobody matches.")

    ready = [c for c in cands if c["status"] == "Hired" and not c["worker_id"]]
    st.markdown("#### Put a hired person in the worker book")
    if not ready:
        st.caption("Nobody is waiting. Set a candidate to 'Hired' above and press Save, then come back here.")
    else:
        names = {f"{c['name_en'] or c['name_kh']} ({c['position'] or '-'})": c for c in ready}
        pick = st.selectbox("Who", list(names), key="hr_hire_pick")
        c = names[pick]
        old = H.rehire_check(c["id_card"], exclude_candidate=c["id"])
        if old:
            st.warning("⚠️ This ID card number is already known:\n\n" + "\n\n".join("- " + x for x in old) +
                       "\n\nIf this is a returning worker, check the old profile first. You can still give a new ID below.")
        h = st.columns(4)
        wid = h[0].text_input("New worker ID", value=H.next_worker_id(), key=f"hr_hire_id_{c['id']}")
        dept = h[1].text_input("Department", value=c["dept"] or "", key=f"hr_hire_d_{c['id']}")
        title = h[2].text_input("Position", value=c["position"] or "", key=f"hr_hire_t_{c['id']}")
        entry = h[3].date_input("Entry date", value=H._d(c["start_on"]) or date.today(), min_value=MIN_DATE, max_value=MAX_DATE, key=f"hr_hire_e_{c['id']}")
        st.caption("This makes a new worker with this ID and a first 3-month contract (not made yet) and an empty document list.")
        if st.button("Add to the worker book", type="primary", key=f"hr_hire_go_{c['id']}"):
            err = H.hire_to_worker(c["id"], wid, dept, title, _iso(entry))
            if err:
                st.error(err)
            else:
                st.session_state["hr_flash"] = f"{wid.strip().upper()} is now in the worker book. Open HR > Workers to add the photo and documents."
                st.rerun()

    st.markdown("#### People you recruited")
    rec = [c for c in cands if c["status"] == "Hired" or c["worker_id"]]
    if rec:
        df = pd.DataFrame([{"Name": c["name_en"] or c["name_kh"], "Phone": c["phone"], "For": f"{c['position'] or '-'} ({c['dept'] or '-'})",
                            "Found on": c["found_on"], "Starts on": c["start_on"], "Worker ID": c["worker_id"] or "(not in book yet)"} for c in rec])
        st.dataframe(df, hide_index=True, width="stretch", height=min(60 + 35 * len(df), 420))
        st.download_button("⬇️ Download (opens in Excel)", _csv(df), file_name="recruited.csv", mime="text/csv")
    else:
        st.caption("Nobody hired yet.")


# ================================================================== RESIGNED
def render_resigned():
    _head("Resigned workers", "Four steps for each worker who left: resign paper, final pay, items returned, NSSF stop. Type a date when a step is done.")
    rows = H.resigned_rows()
    today = date.today()
    tracked = [r for r in rows if r["tracked"]]
    k = st.columns(4)
    k[0].metric("Resigned this month", sum(1 for r in rows if r["stop_date"][:7] == f"{today:%Y-%m}"))
    k[1].metric("Not finished", sum(r["open"] for r in rows))
    k[2].metric("Not paid yet", sum(1 for r in tracked if not r["paid"]))
    k[3].metric("NSSF stop not reported", sum(1 for r in tracked if not r["nssf_reported"]))

    with st.expander("➕ A worker is leaving / left (set the last working day)"):
        opts = _worker_options(True)
        if not opts:
            st.info("No working worker in the book.")
        else:
            c = st.columns([3, 1, 2])
            who = c[0].selectbox("Worker", opts, key="hr_rs_who")
            day = c[1].date_input("Last working day", value=today, min_value=MIN_DATE, max_value=MAX_DATE, key="hr_rs_day")
            why = c[2].text_input("Reason", key="hr_rs_why")
            if st.button("Mark as resigned", type="primary", key="hr_rs_go"):
                err = H.mark_resigned(_wid_of(who), _iso(day), why)
                st.session_state["hr_flash"] = err or f"{_wid_of(who)} is marked as resigned. Fill the four steps below."
                st.rerun()

    show = st.radio("Show", ["Needs action", "All tracked", "Older (before I started tracking)"], horizontal=True, key="hr_rs_show")
    if show == "Needs action":
        lst = [r for r in rows if r["open"]]
    elif show == "All tracked":
        lst = tracked
    else:
        lst = [r for r in rows if not r["tracked"]]
    f1, f2 = st.columns([1, 2])
    depts = sorted({r["dept"] for r in lst if r["dept"]})
    dep = f1.selectbox("Department", ["All departments"] + depts, key="hr_rs_dep")
    q = f2.text_input("Search", placeholder="Name or worker ID", key="hr_rs_q").strip().lower()
    lst = [r for r in lst if (dep == "All departments" or r["dept"] == dep) and (not q or q in r["worker_id"].lower() or q in (r["name"] or "").lower() or q in (r["name_kh"] or ""))]
    if not lst:
        st.success("Nothing in this list.")
    else:
        df = pd.DataFrame([{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Left on": H._d(r["stop_date"]),
                            "Excel text": r["stop_raw"], "Reason": r["reason"], "Resign paper filled on": H._d(r["paper_on"]), "Paid on": H._d(r["paid_on"]),
                            "Items returned on": H._d(r["items_on"]), "NSSF stop reported on": H._d(r["nssf_on"]), "Note": r["note"]} for r in lst])
        ver = st.session_state.get("hr_rsl_ver", 0)
        ed = st.data_editor(df, hide_index=True, width="stretch", height=min(60 + 36 * len(df), 640),
                            key=f"hr_rsl_{ver}_{show}_{len(df)}_{hash(tuple(df['Worker'])) % 99999}",
                            disabled=["Worker", "Name", "Dept", "Excel text"],
                            column_config={"Left on": _DC, "Resign paper filled on": _DC, "Paid on": _DC, "Items returned on": _DC, "NSSF stop reported on": _DC})
        c1, c2 = st.columns([1, 1])
        if c1.button("💾 Save", type="primary", key="hr_rsl_save"):
            errs, n = [], 0
            for r, (_, e) in zip(lst, ed.iterrows()):
                po, pd_, io, no = _tx(e["Resign paper filled on"]), _tx(e["Paid on"]), _tx(e["Items returned on"]), _tx(e["NSSF stop reported on"])
                reason, note = str(e["Reason"] or "").strip(), str(e["Note"] or "").strip()
                left_on = _tx(e["Left on"])
                if not (po or pd_ or io or no or reason or note or r["record"]) and left_on == (r["stop_date"] or ""):
                    continue
                if left_on != (r["stop_date"] or "") and left_on:
                    err = H.mark_resigned(r["worker_id"], left_on, reason)
                    if err:
                        errs.append(err)
                        continue
                err = H.save_resign({"worker_id": r["worker_id"], "reason": reason, "paper_done": bool(po), "paper_on": po, "paid": bool(pd_), "paid_on": pd_,
                                     "items_back": bool(io), "items_on": io, "nssf_reported": bool(no), "nssf_on": no, "note": note})
                if err:
                    errs.append(err)
                else:
                    n += 1
            if errs:
                st.error("Some lines were not saved:\n\n" + "\n\n".join(errs))
            if n:
                st.session_state["hr_rsl_ver"] = ver + 1
                st.session_state["hr_flash"] = f"Saved {n} line(s)."
                st.rerun()
        with st.expander(f"⚡ Quick fill: one date for all {len(lst)} workers shown"):
            st.caption("Use this once to catch up (example: all of them were already paid on the same day). Only workers that do not have this step yet are changed.")
            q1, q2, q3 = st.columns([2, 1, 1])
            step = q1.selectbox("Step", ["Resign paper filled", "Paid", "Items returned", "NSSF stop reported"], key="hr_rs_step")
            sday = q2.date_input("Date", value=today, min_value=MIN_DATE, max_value=MAX_DATE, key="hr_rs_stepd")
            fk = {"Resign paper filled": ("paper_done", "paper_on"), "Paid": ("paid", "paid_on"),
                  "Items returned": ("items_back", "items_on"), "NSSF stop reported": ("nssf_reported", "nssf_on")}[step]
            todo = [r for r in lst if not r[fk[0]]]
            if q3.button(f"Fill {len(todo)} worker(s)", key="hr_rs_fill", disabled=not todo):
                for r in todo:
                    x = {k: r[k] for k in ("reason", "paper_done", "paper_on", "paid", "paid_on", "items_back", "items_on", "nssf_reported", "nssf_on", "note")}
                    x.update({"worker_id": r["worker_id"], fk[0]: 1, fk[1]: _iso(sday)})
                    H.save_resign(x)
                st.session_state["hr_rsl_ver"] = ver + 1
                st.session_state["hr_flash"] = f"'{step}' filled for {len(todo)} worker(s)."
                st.rerun()
        out = df.copy()
        for c in ("Left on", "Resign paper filled on", "Paid on", "Items returned on", "NSSF stop reported on"):
            out[c] = out[c].map(lambda v: "" if v is None or v != v else f"{v:%Y-%m-%d}")
        c2.download_button("⬇️ Download (opens in Excel)", _csv(out), file_name="resigned.csv", mime="text/csv")

    old = [r for r in rows if not r["tracked"]]
    if old:
        with st.expander(f"Older resigned workers ({len(old)}) that I never tracked"):
            st.caption(f"These left more than {H.get_setting('track_resign_days')} days ago (or have no readable date) and have no record, so they are not counted as 'not paid'. "
                       "If they are all finished, you can close them with one button. You can still open each one and change it.")
            dated = [r for r in old if r["stop_date"]]
            if dated and st.button(f"Mark the {len(dated)} older workers as all finished", key="hr_rs_closeold"):
                n = H.mark_old_resigned_done(today.isoformat())
                st.session_state["hr_flash"] = f"{n} older resigned workers closed."
                st.rerun()


# ================================================================== REMINDERS + WARNING LETTERS
def render_reminders():
    _head("Reminders and warnings", "Things to remember (including maternity), birthdays, and warning letters.")
    t1, t2, t3 = st.tabs(["🔔 Reminders", "🎂 Birthdays", "⚠️ Warning letters"])
    with t1:
        with st.form("hr_new_rem", clear_on_submit=True):
            a = st.columns([1, 1, 2])
            due = a[0].date_input("Date", value=date.today(), min_value=MIN_DATE, max_value=MAX_DATE)
            kind = a[1].selectbox("Kind", H.REMINDER_KINDS)
            wid = a[2].text_input("Worker ID (optional)", placeholder="DMO-10002")
            text = st.text_input("What to remember (example: Maternity leave ends, call her)")
            if st.form_submit_button("Add reminder", type="primary"):
                err = H.add_reminder(_iso(due), kind, wid, text)
                if err:
                    st.error(err)
                else:
                    st.session_state["hr_flash"] = "Reminder saved."
                    st.rerun()
        rem = H.list_reminders(only_open=True)
        if not rem:
            st.info("Nothing to remember right now.")
        else:
            df = pd.DataFrame([{"Date": H._d(r["due"]), "Kind": r["kind"], "Worker": r["worker_id"], "What": r["text"],
                                "When": ("OVERDUE" if H._d(r["due"]) < date.today() else ("Today" if H._d(r["due"]) == date.today() else "")), "Done": False, "Delete": False} for r in rem])
            ver = st.session_state.get("hr_rm_ver", 0)
            ed = st.data_editor(df, hide_index=True, width="stretch", key=f"hr_rm_{ver}_{len(df)}", disabled=["Date", "Kind", "Worker", "What", "When"],
                                column_config={"Date": _DC, "Done": st.column_config.CheckboxColumn(width=60), "Delete": st.column_config.CheckboxColumn(width=60)})
            if st.button("💾 Save", type="primary", key="hr_rm_save"):
                H.set_reminders_done([r["id"] for r, (_, e) in zip(rem, ed.iterrows()) if e["Done"] and not e["Delete"]],
                                     [r["id"] for r, (_, e) in zip(rem, ed.iterrows()) if e["Delete"]])
                st.session_state["hr_rm_ver"] = ver + 1
                st.session_state["hr_flash"] = "Saved."
                st.rerun()
    with t2:
        span = st.radio("Show birthdays in the next", [7, 30, 60], horizontal=True, format_func=lambda x: f"{x} days", key="hr_bd_span")
        b = H.birthdays(span)
        if not b:
            st.info("No birthday in this time (or the birthdays are empty in the worker book).")
        else:
            df = pd.DataFrame([{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Birthday": r["birthday"], "In days": r["in_days"], "Turns": r["turns"]} for r in b])
            st.dataframe(df, hide_index=True, width="stretch")
            st.download_button("⬇️ Download (opens in Excel)", _csv(df), file_name="birthdays.csv", mime="text/csv")
    with t3:
        opts = _worker_options(False)
        if opts:
            c = st.columns([3, 1, 1, 2])
            who = c[0].selectbox("Worker", opts, key="hr_wl_who")
            on = c[1].date_input("Date", value=date.today(), min_value=MIN_DATE, max_value=MAX_DATE, key="hr_wl_d")
            kind = c[2].selectbox("Kind", H.WARNING_KINDS, key="hr_wl_k")
            why = c[3].text_input("Reason", key="hr_wl_r")
            if st.button("Add warning letter", type="primary", key="hr_wl_add"):
                err = H.add_warning(_wid_of(who), _iso(on), kind, why)
                st.session_state["hr_flash"] = err or "Warning letter saved."
                st.rerun()
        ws = H.list_warnings()
        q = st.text_input("Search the warning letters", placeholder="Name or worker ID", key="hr_wl_q").strip().lower()
        ws = [w for w in ws if not q or q in w["worker_id"].lower() or q in (w.get("name_en") or "").lower()]
        if not ws:
            st.info("No warning letters.")
        else:
            counts = {}
            for w in H.list_warnings():
                counts[w["worker_id"]] = counts.get(w["worker_id"], 0) + 1
            df = pd.DataFrame([{"Date": w["on_date"], "Worker": w["worker_id"], "Name": w.get("name_en") or "", "Dept": w.get("dept") or "", "Kind": w["kind"],
                                "Reason": w["reason"], "Letters of this worker": counts[w["worker_id"]]} for w in ws])
            st.dataframe(df, hide_index=True, width="stretch", height=min(60 + 35 * len(df), 460))
            st.download_button("⬇️ Download (opens in Excel)", _csv(df), file_name="warning_letters.csv", mime="text/csv")


# ================================================================== REPORTS
def render_reports():
    _head("Reports", "Monthly numbers, NSSF lists, leave count from your codes, and the weekly summary for the boss.")
    t1, t2, t3, t4, t5 = st.tabs(["📊 Monthly report", "🏛️ NSSF", "🗓️ Leave count", "📨 Boss summary", "💵 Holiday work (paid on the day)"])
    today = date.today()
    with t1:
        c = st.columns(2)
        yr = c[0].number_input("Year", 2020, 2100, today.year, key="hr_rp_y")
        mo = c[1].selectbox("Month", list(range(1, 13)), index=today.month - 1, format_func=lambda m: date(2000, m, 1).strftime("%B"), key="hr_rp_m")
        m = H.month_report(int(yr), int(mo))
        k = st.columns(5)
        k[0].metric("At the start", m["start"])
        k[1].metric("Hired", m["hired"])
        k[2].metric("Resigned", m["left"])
        k[3].metric("At the end", m["end"])
        k[4].metric("Women / Men at the end", f"{m['women']} / {m['men']}")
        if m["no_entry"]:
            st.warning(f"{m['no_entry']} worker(s) have no entry date. They are counted as already working before this month.")
        df = pd.DataFrame([{"Department": r["dept"], "At the start": r["start"], "Hired": r["hired"], "Resigned": r["left"], "At the end": r["end"], "Women at the end": r["end_women"]} for r in m["rows"]])
        st.dataframe(df, hide_index=True, width="stretch")
        st.download_button("⬇️ Download this month (opens in Excel)", _csv(df), file_name=f"monthly_report_{int(yr)}-{int(mo):02d}.csv", mime="text/csv")
        st.caption("This is my own table. If the Ministry (MLVT) gives you a form with exact columns, send it to me and I will make the report match it.")
        a, b = st.columns(2)
        with a:
            st.markdown("**Hired this month**")
            st.dataframe(pd.DataFrame([{"ID": w["id"], "Name": w["name_en"], "Dept": w["dept"], "Entry": w["entry_date"]} for w in m["hired_list"]]), hide_index=True, width="stretch")
        with b:
            st.markdown("**Resigned this month**")
            st.dataframe(pd.DataFrame([{"ID": w["id"], "Name": w["name_en"], "Dept": w["dept"], "Left": w["stop_date"]} for w in m["left_list"]]), hide_index=True, width="stretch")
        if st.button(f"💾 Save the headcount of {date(int(yr), int(mo), 1):%B %Y} (month snapshot)", key="hr_rp_snap"):
            H.save_snapshot(int(yr), int(mo))
            st.session_state["hr_flash"] = "Snapshot saved."
            st.rerun()
        snaps = H.list_snapshots()
        if snaps:
            st.markdown("**Saved month snapshots**")
            st.dataframe(pd.DataFrame([{"Month": s["month"], "Working": s["working"], "Women": s["women"], "Men": s["men"], "Hired": s["hired"], "Resigned": s["resigned"], "Saved": s["saved_at"]} for s in snaps]),
                         hide_index=True, width="stretch")
            st.line_chart(pd.DataFrame([{"Month": s["month"], "Working": s["working"]} for s in reversed(snaps)]).set_index("Month"), height=220, color=["#1F6F5F"])
    with t2:
        ws = H.list_workers()
        no_n = [w for w in ws if w["status"] == "Working" and not w["nssf"]]
        st.markdown(f"**Working without an NSSF number: {len(no_n)}**")
        df = pd.DataFrame([{"ID": w["id"], "Name": w["name_en"], "Khmer": w["name_kh"], "Dept": w["dept"], "Entry": w["entry_date"], "ID card": w["id_card"]} for w in no_n])
        st.dataframe(df, hide_index=True, width="stretch", height=min(60 + 35 * max(len(df), 1), 380))
        if len(df):
            st.download_button("⬇️ Download (opens in Excel)", _csv(df), file_name="no_nssf_number.csv", mime="text/csv")
        rr = [r for r in H.resigned_rows() if r["tracked"] and not r["nssf_reported"]]
        st.markdown(f"**Resigned, NSSF stop not reported: {len(rr)}**")
        st.dataframe(pd.DataFrame([{"ID": r["worker_id"], "Name": r["name"], "NSSF": r["nssf"], "Left on": r["stop_date"]} for r in rr]), hide_index=True, width="stretch")
        st.caption("You mark 'NSSF stop reported' in HR > Resigned.")
    with t3:
        st.caption("Counts the codes you saved in Daily work (Step 3 decisions). Only days you checked and saved are counted.")
        c = st.columns(3)
        yr = c[0].number_input("Year", 2020, 2100, today.year, key="hr_lc_y")
        mo = c[1].selectbox("Month", list(range(1, 13)), index=today.month - 1, format_func=lambda m: date(2000, m, 1).strftime("%B"), key="hr_lc_m")
        q = c[2].text_input("Search", placeholder="Name or worker ID", key="hr_lc_q").strip().lower()
        table, names = H.codes_in_month(int(yr), int(mo))
        codes = sorted({k for t in table.values() for k in t})
        rows = [{"ID": wid, "Name": names.get(wid, ("", ""))[0], "Dept": names.get(wid, ("", ""))[1], **{k: t.get(k, 0) for k in codes}} for wid, t in sorted(table.items())
                if not q or q in wid.lower() or q in names.get(wid, ("", ""))[0].lower()]
        if not rows:
            st.info("No saved codes for this month.")
        else:
            df = pd.DataFrame(rows)
            st.dataframe(df, hide_index=True, width="stretch", height=min(60 + 35 * len(df), 520))
            st.download_button("⬇️ Download (opens in Excel)", _csv(df), file_name=f"codes_{int(yr)}-{int(mo):02d}.csv", mime="text/csv")
    with t4:
        txt = H.boss_summary()
        st.caption("Copy this and send it to the boss (the copy icon is at the top right of the box).")
        st.code(txt, language=None)
        st.download_button("⬇️ Download as a text file", txt.encode("utf-8"), file_name=f"hr_weekly_{today:%Y-%m-%d}.txt", mime="text/plain")


    with t5:
        holiday_paid_tab()


def holiday_paid_tab():
    st.caption("Workers who worked on a holiday (or a bridge day) and were PAID ON THE DAY. This list is what you take out of the payroll of that month. "
               "Rate: holiday = hours x 2 dollars, normal day = hours x 1 dollar (you can change it below).")
    flash = st.session_state.pop("hr_hp_flash", "")
    if flash:
        st.success(flash)
    months = H.hp_months()
    today = date.today()
    cur = f"{today:%Y-%m}"
    if cur not in months:
        months = [cur] + months
    month = st.selectbox("Month", months, key="hr_hp_month")
    rows = H.hp_list(month)
    k = st.columns(4)
    k[0].metric("Lines", len(rows))
    k[1].metric("Workers", len({r["worker_id"] for r in rows}))
    k[2].metric("Hours", f"{sum(r['hours'] for r in rows):g}")
    k[3].metric("Paid already (USD)", f"{sum(r['amount'] for r in rows):g}")
    if rows:
        df = pd.DataFrame([{"Worker": r["worker_id"], "Name": r["name"], "Dept": r["dept"], "Date": _d(r["day"]), "Holiday": r["holiday_name"], "Hours": r["hours"],
                            "Amount": r["amount"], "Date paid": _d(r["date_paid"]), "Method": r["method"], "Exclude": r["exclude"], "Delete": False} for r in rows])
        ver = st.session_state.get("hr_hp_ver", 0)
        ed = st.data_editor(df, hide_index=True, width="stretch", height=min(60 + 36 * len(df), 520), key=f"hr_hp_ed_{month}_{ver}_{len(df)}",
                            disabled=["Worker", "Name", "Dept", "Date"],
                            column_config={"Date": st.column_config.DateColumn(format="DD MMM YYYY"),
                                           "Date paid": st.column_config.DateColumn(format="DD MMM YYYY", min_value=MIN_DATE, max_value=MAX_DATE),
                                           "Exclude": st.column_config.SelectboxColumn(options=["YES", "NO"], width=80),
                                           "Delete": st.column_config.CheckboxColumn(width=60)})
        b1, b2 = st.columns([1, 1])
        with b1:
            if st.button("💾 Save my changes", type="primary", key="hr_hp_save"):
                errs, dele = [], []
                for r, (_, e) in zip(rows, ed.iterrows()):
                    if e["Delete"]:
                        dele.append(r["id"])
                        continue
                    err = H.hp_save({**r, "holiday_name": e["Holiday"], "hours": e["Hours"], "amount": e["Amount"], "date_paid": _tx(e["Date paid"]),
                                     "method": e["Method"], "exclude": e["Exclude"]})
                    if err:
                        errs.append(err)
                if dele:
                    H.hp_delete(dele)
                if errs:
                    st.error("Some lines were not saved:\n\n" + "\n\n".join(errs))
                else:
                    st.session_state["hr_hp_ver"] = ver + 1
                    st.session_state["hr_hp_flash"] = "Saved."
                    st.rerun()
        with b2:
            st.download_button("⬇️ Download the Excel (same as your list)", H.hp_export_xlsx(month), file_name=f"Holiday_Work_List_{month}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="hr_hp_dl")
    else:
        st.info("No line for this month yet.")
    with st.expander("➕ Add workers who worked on a day", expanded=not rows):
        d1, d2, d3 = st.columns([1, 1.3, 1])
        day = d1.date_input("Date", value=today, key="hr_hp_day", format="DD/MM/YYYY")
        hol_names = {}
        try:
            import attendance_agent as _A
            hol_names = _A.load_holidays()
        except Exception:
            pass
        hname = d2.text_input("Holiday name (NORMAL DAY if it is a normal day)", value=hol_names.get(str(day), "NORMAL DAY"), key=f"hr_hp_hn_{day}")
        method = d3.text_input("Paid by", value="ABA", key="hr_hp_method")
        rate = H.hp_rate(hname)
        saved = {}
        try:
            saved = {x["worker_id"]: x for x in db.day_decisions(day)}
        except Exception:
            pass
        base = []
        try:
            ws = {w["id"]: w for w in H.list_workers()}
            for r in db.day_records(day):
                if r["in_time"]:
                    h = saved.get(r["worker_id"], {}).get("hours") or r["paper_hours"] or "8"
                    try:
                        h = float(h)
                    except ValueError:
                        h = 8.0
                    base.append({"Worker": r["worker_id"], "Name": r["name"], "Hours": h})
        except Exception:
            base = []
        st.caption(f"{len(base)} worker(s) had fingerprint scans on that day in your saved check (their hours come from your decision or the paper). "
                   f"Remove the ones who were not paid on the day. You can add a worker with the + at the bottom. Amount = hours x {rate:g}.")
        tbl = pd.DataFrame(base or [{"Worker": "", "Name": "", "Hours": 8.0}])
        ed2 = st.data_editor(tbl, num_rows="dynamic", hide_index=True, width="stretch", key=f"hr_hp_new_{day}_{len(base)}",
                             column_config={"Hours": st.column_config.NumberColumn(min_value=0.5, max_value=24, step=0.5)})
        if st.button("Add these lines to the list", key="hr_hp_add", type="primary"):
            ws = {w["id"]: w for w in H.list_workers()}
            n, errs = 0, []
            for _, e in ed2.iterrows():
                wid = str(e["Worker"] or "").strip().upper()
                if not wid:
                    continue
                w = ws.get(wid, {})
                hrs = float(e["Hours"] or 0)
                err = H.hp_save({"worker_id": wid, "name": w.get("name_en", "") or e["Name"], "dept": w.get("dept", ""), "position": w.get("title", ""), "day": day.isoformat(),
                                 "holiday_name": hname, "hours": hrs, "amount": hrs * rate, "date_paid": day.isoformat(), "method": method, "exclude": "YES"})
                if err:
                    errs.append(err)
                else:
                    n += 1
            if errs:
                st.error("Some lines were not added:\n\n" + "\n\n".join(errs))
            if n:
                st.session_state["hr_hp_flash"] = f"Added {n} line(s)."
                st.rerun()
    with st.expander("📥 Bring in my old Holiday Work List Excel"):
        up = st.file_uploader("Your Excel (sheet 'Holiday Work List')", type=["xlsx"], key="hr_hp_up")
        if up is not None and st.button("Import this file", key="hr_hp_imp"):
            if len(up.getvalue()) > 10 * 1024 * 1024:
                st.error("The file is bigger than 10 MB.")
            else:
                n, errs = H.hp_import_xlsx(up.getvalue())
                if errs:
                    st.error("\n\n".join(errs[:8]))
                if n:
                    st.session_state["hr_hp_flash"] = f"Imported {n} line(s). A worker + date that was already in the list was replaced."
                    st.rerun()
    with st.expander("⚙️ Rates and contract pay"):
        r1, r2 = st.columns(2)
        h2 = r1.number_input("Holiday: dollars per hour", 0, 50, H.get_setting("hp_rate_holiday"), key="hr_hp_r1")
        h1 = r2.number_input("Normal day: dollars per hour", 0, 50, H.get_setting("hp_rate_normal"), key="hr_hp_r2")
        if st.button("Save the rates", key="hr_hp_rates"):
            H.set_setting("hp_rate_holiday", h2); H.set_setting("hp_rate_normal", h1)
            st.session_state["hr_hp_flash"] = "Rates saved."
            st.rerun()


# ================================================================== FACTORY DOCUMENTS
def render_factory():
    _head("Factory documents", "Licenses, insurance, certificates. Put the expiry date and I will warn you before it ends.")
    with st.form("hr_fd_new", clear_on_submit=True):
        a = st.columns([2, 1, 1])
        name = a[0].text_input("Name of the document")
        kind = a[1].selectbox("Kind", H.FACTORY_KINDS)
        exp = a[2].date_input("Expiry date (leave empty if none)", value=None, min_value=MIN_DATE, max_value=MAX_DATE)
        note = st.text_input("Note")
        up = st.file_uploader("File (optional): PDF, JPG, PNG, Word, Excel, up to 10 MB")
        if st.form_submit_button("Save this document", type="primary"):
            err = H.add_factory_doc(name, kind, _iso(exp), note, up.name if up else None, up.getvalue() if up else None)
            if err:
                st.error(err)
            else:
                st.session_state["hr_flash"] = "Document saved."
                st.rerun()
    docs = H.list_factory_docs()
    if not docs:
        st.info("No factory document yet.")
        return
    for d in docs:
        a, b, c, e = st.columns([3, 1.4, 1.2, 1])
        a.markdown(f"**{html.escape(d['name'])}**<br><span class='hr-sub'>{html.escape(d['kind'])} {('· ' + html.escape(d['note'])) if d['note'] else ''}</span>", unsafe_allow_html=True)
        if d["days"] is None:
            b.markdown("<span class='hr-pill hr-green'>No expiry</span>", unsafe_allow_html=True)
        elif d["days"] < 0:
            b.markdown(f"<span class='hr-pill hr-red'>Expired {-d['days']} days ago</span>", unsafe_allow_html=True)
        elif d["days"] <= 30:
            b.markdown(f"<span class='hr-pill hr-amber'>Expires in {d['days']} days</span>", unsafe_allow_html=True)
        else:
            b.markdown(f"<span class='hr-pill hr-green'>Until {d['expiry']}</span>", unsafe_allow_html=True)
        if d["path"]:
            with open(d["path"], "rb") as fh:
                c.download_button("Open / download", fh.read(), file_name=d["file"].split("__", 1)[-1], key=f"hr_fd_dl_{d['id']}")
        if e.button("Delete", key=f"hr_fd_del_{d['id']}"):
            H.delete_factory_doc(d["id"])
            st.rerun()
