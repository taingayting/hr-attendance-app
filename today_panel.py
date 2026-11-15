"""The 'Today' checklist at the top of Daily work: what is open, so nothing is forgotten."""
from datetime import date

import streamlit as st

import attendance_db as db


def _go_day(d):
    st.session_state["work_day"] = d


def _go_page(p):
    st.session_state["page_att"] = p


def skipped_days(fp_days, A, today=None):
    """Working days that are in the fingerprint file, before the last day you filled, and not filled yet (only the same month as the last filled day)."""
    lf = db.last_filled_day()
    if not lf or not fp_days:
        return []
    last = date.fromisoformat(lf)
    done = {r["day"] for r in db.days_list() if r.get("filled_at")}
    out = []
    for d in sorted(fp_days):
        if d >= last or (d.year, d.month) != (last.year, last.month):
            continue
        try:
            off = A.day_type(d) != "normal"
        except Exception:
            off = d.weekday() == 6
        if not off and d.isoformat() not in done:
            out.append(d)
    return out


def render(slot, day, fp_days, steps, A):
    """steps = (files_ok, paper_ok, checked, saved)"""
    today = date.today()
    try:
        notes = db.notes_due(today)
        cases = db.case_overdue(today)
        letters = db.letters_list()
        not_told = [x for x in letters if not x.get("told_on") and not x.get("returned_on")]
        not_back = [x for x in letters if x.get("told_on") and not x.get("returned_on")]
        skipped = skipped_days(fp_days, A)
    except Exception as e:
        with slot:
            st.caption(f"Today list not available: {e}")
        return
    names = ["1. Files", "2. Paper list", "3. Check + Excel report", "4. Fill your Excel"]
    open_steps = [n for n, ok in zip(names, steps) if not ok]
    n_open = len(notes) + len(cases) + len(not_told) + len(not_back) + len(skipped) + len(open_steps)
    with slot:
        with st.expander(f"📋 Today · {'all clear ✅' if not n_open else f'{n_open} thing(s) open'}", expanded=bool(n_open)):
            def line(ok, text, page=None, key=None):
                c = st.columns([5, 1])
                c[0].markdown(f"{'✅' if ok else '🔴'} {text}")
                if page and not ok:
                    c[1].button("Open", key=key, on_click=_go_page, args=(page,), width="stretch")
            line(not notes, f"Reminders due: **{len(notes)}**" + (" (red box above)" if notes else ""), "Notes", "td_notes")
            line(not cases, f"Questions nobody asked for 2+ days: **{len(cases)}**", "Questions to ask", "td_cases")
            line(not not_told, f"Missed-scan letters not yet told to the worker: **{len(not_told)}**", "Letters", "td_l1")
            line(not not_back, f"Letters told but not returned yet: **{len(not_back)}**", "Letters", "td_l2")
            line(not open_steps, f"This day ({day:%d %b}): " + (", ".join(f"⬜ {n}" for n in open_steps) if open_steps else "all 4 steps done"))
            if skipped:
                st.markdown(f"🔴 Working days with fingerprint data, before your last filled day, **not filled**: **{len(skipped)}**")
                cols = st.columns(min(len(skipped), 6))
                for i, d in enumerate(skipped[:6]):
                    cols[i].button(f"{d:%d %b}", key=f"td_d_{d}", on_click=_go_day, args=(d,), width="stretch")
                if len(skipped) > 6:
                    st.caption("…and " + ", ".join(f"{d:%d %b}" for d in skipped[6:]))
            else:
                st.markdown("✅ No skipped working days.")
