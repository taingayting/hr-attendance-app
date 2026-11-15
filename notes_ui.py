"""Quick notes and reminders: a box in the sidebar on every page, a red banner when something is due, a Notes page, and a history per worker."""
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import attendance_db as db

WHEN = ["No date (just a note)", "Tomorrow", "In 2 days", "Next week", "Pick a date"]


def _sync():
    try:
        import cloud_sync
        if cloud_sync.enabled():
            cloud_sync.push()
    except Exception:
        pass


def _add(text, wid, due):
    err = db.note_add(text, wid, due)
    st.session_state["note_msg"] = err or "✅ Saved." + (f" I will remind you on {due}." if due else "")
    if not err:
        st.session_state["note_ver"] = st.session_state.get("note_ver", 0) + 1
    _sync()


def sidebar_box():
    with st.sidebar.expander("📝 Note / remind me", expanded=False):
        v = st.session_state.get("note_ver", 0)
        text = st.text_area("What do you want to remember?", key=f"nt_text_{v}", height=80, placeholder="Example: Leader said DMO-10078 was sick, ask again Friday")
        wid = st.text_input("Worker ID (optional)", key=f"nt_wid_{v}", placeholder="DMO-10078")
        when = st.selectbox("Remind me", WHEN, key=f"nt_when_{v}")
        today = date.today()
        due = ""
        if when == "Tomorrow":
            due = (today + timedelta(days=1)).isoformat()
        elif when == "In 2 days":
            due = (today + timedelta(days=2)).isoformat()
        elif when == "Next week":
            due = (today + timedelta(days=7)).isoformat()
        elif when == "Pick a date":
            due = st.date_input("Date", value=today + timedelta(days=1), min_value=today, key=f"nt_date_{v}").isoformat()
        st.button("💾 Save note", key=f"nt_save_{v}", type="primary", width="stretch", on_click=_add, args=(text, wid, due))
        if st.session_state.get("note_msg"):
            st.caption(st.session_state.pop("note_msg"))


def _done(nid):
    db.note_done(nid)
    _sync()


def _snooze(nid, days):
    db.note_snooze(nid, (date.today() + timedelta(days=days)).isoformat())
    _sync()


def banner():
    """Red box at the top of every page when a reminder is due."""
    try:
        due = db.notes_due(date.today())
    except Exception:
        return
    if not due:
        return
    with st.container(border=True):
        st.markdown(f"#### 🔔 You asked me to remind you ({len(due)})")
        for n in due[:8]:
            c = st.columns([6, 1, 1])
            late = n["due"] < date.today().isoformat()
            who = f" · **{n['worker_id']}**" if n["worker_id"] else ""
            c[0].markdown(f"{'🔴 OVERDUE' if late else '🟠 Today'} ({n['due']}){who}: {n['text']}")
            c[1].button("✅ Done", key=f"nb_done_{n['id']}", on_click=_done, args=(n["id"],), width="stretch")
            c[2].button("⏰ +1 day", key=f"nb_snz_{n['id']}", on_click=_snooze, args=(n["id"], 1), width="stretch")
        if len(due) > 8:
            st.caption(f"…and {len(due) - 8} more on the Notes page.")


def _plain(s):
    import re
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", str(s or ""))).strip()


def worker_history(wid, title="Notes about this worker"):
    try:
        pgs = db.pages_for_worker(wid)
    except Exception:
        pgs = []
    if pgs:
        st.markdown("**📓 Notebook pages about this worker**")
        for p in pgs:
            st.markdown(f"- **{p['title'] or 'Untitled'}** ({p['updated_at'][:10]}) — {_plain(p['snippet'])}")
    rows = db.notes_for_worker(wid)
    if not rows:
        return
    st.markdown(f"**📝 {title}**")
    st.dataframe(pd.DataFrame([{"Written": r["created_at"][:10], "Remind": r["due"], "Done": r["done_at"][:10], "Note": r["text"]} for r in rows]), hide_index=True, width="stretch")


def render():
    st.title("📝 Notes")
    st.caption("Everything you wrote with the 📝 box in the sidebar. A note with a date will pop up in a red box on every page when the day comes. Tick Done when finished.")
    c = st.columns([1, 1, 2])
    state = c[0].radio("Show", ["open", "done", "all"], horizontal=True, key="nt_state", format_func=lambda x: {"open": "Not done", "done": "Done", "all": "All"}[x])
    wid = c[1].text_input("Worker ID", key="nt_filter_wid", placeholder="DMO-10078")
    q = c[2].text_input("Search the words", key="nt_q")
    rows = db.notes_list(state, wid, q)
    if not rows:
        st.info("No notes here yet. Use the 📝 box in the sidebar.")
        return
    today = date.today().isoformat()
    df = pd.DataFrame([{"Written": r["created_at"][:10], "Worker": r["worker_id"], "Remind on": r["due"],
                        "When": ("DONE" if r["done_at"] else "OVERDUE" if r["due"] and r["due"] < today else "Today" if r["due"] == today else ""),
                        "Note": r["text"], "Done": bool(r["done_at"]), "Delete": False} for r in rows])
    ver = st.session_state.get("nt_ver", 0)
    ed = st.data_editor(df, hide_index=True, width="stretch", key=f"nt_ed_{ver}_{len(df)}_{state}", disabled=["Written", "Worker", "Remind on", "When", "Note"],
                        column_config={"Done": st.column_config.CheckboxColumn(width=60), "Delete": st.column_config.CheckboxColumn(width=60)})
    if st.button("💾 Save", type="primary", key="nt_save_page"):
        for r, (_, e) in zip(rows, ed.iterrows()):
            if e["Delete"]:
                db.note_delete(r["id"])
            elif bool(e["Done"]) != bool(r["done_at"]):
                db.note_done(r["id"], bool(e["Done"]))
        st.session_state["nt_ver"] = ver + 1
        _sync()
        st.rerun()
