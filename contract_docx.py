"""
Fills YOUR contract paper (a Word file you upload once) with one worker's data. The Word file is only read: a new file is made.
It finds the places to fill by the printed labels (worker ID, sex, birth date ...), so the layout of your paper stays as it is.
"""
import io
from datetime import date

import docx

NB = " "


def _blank(t):
    return t.replace(NB, "").strip() == ""


def _find(runs, text, start=0, exact=False):
    for i in range(start, len(runs)):
        t = runs[i].text
        if (t.strip() == text) if exact else (text in t):
            return i
    return -1


def _set_span(runs, a, b, value):
    """Between run a and run b (both excluded): the first run with text gets `value`, the other runs with text are emptied."""
    first = True
    for r in runs[a + 1:b]:
        if _blank(r.text):
            continue
        if first:
            r.text, first = value, False
        else:
            r.text = ""
    if first:                                  # nothing was written there: put the value in the first run of the span
        if b > a + 1:
            runs[a + 1].text = value
        else:
            raise ValueError("empty place")


def _digits(runs, a, b):
    return [r for r in runs[a + 1:b] if r.text.replace(NB, "").strip().isdigit()]


def _para(d, text):
    for p in d.paragraphs:
        if text in p.text:
            return p
    raise ValueError(f"I did not find '{text}' in your contract paper. Is it the usual paper?")


def fill(template_bytes, w, c, salary, allowance):
    """w = worker dict from the worker book, c = dict(start, end) as dates. Returns the new .docx bytes. Raises ValueError with a plain message."""
    d = docx.Document(io.BytesIO(template_bytes))
    name = (w.get("name_kh") or w.get("name_en") or "").strip()
    sd, ed = c["start"], c["end"]
    # ----- the line with the worker's data
    runs = _para(d, "工号").runs
    pos = 0
    plan = [("姓名", "អត្ថលេខ", name), ("工号", "ភេទ", w["id"]), ("性别", "សញ្ញាតិ", w.get("sex", "")),
            ("出生日期", "លេខអត្តសញ្ញាណប័ណ្ណ", w.get("birthday", "")), ("身份证号码", "ទីកន្លែងកំណើត", w.get("id_card", "")),
            ("出生址", "មុខងារ", w.get("province", ""))]
    for lab, nxt, val in plan:
        a = _find(runs, lab, pos, exact=True)
        b = _find(runs, nxt, a + 1, exact=True)
        if a < 0 or b < 0:
            raise ValueError(f"I could not find the place for '{lab}' in your contract paper.")
        if val:
            _set_span(runs, a, b, val)
        pos = a
    a = _find(runs, "មុខងារ", pos, exact=True)
    b = _find(runs, "កាលបរិច្ឆេទចូលបំរើការងារ", a + 1)
    if a < 0 or b < 0:
        raise ValueError("I could not find the place for the position in your contract paper.")
    if w.get("title"):
        _set_span(runs, a, b, w["title"])
    a = _find(runs, "入职日期", b, exact=True)
    b = _find(runs, "អាស័យដ្ឋានបច្ចុប្បន្ន", a + 1, exact=True)
    if a < 0 or b < 0:
        raise ValueError("I could not find the place for the entry date in your contract paper.")
    if w.get("entry_date"):
        _set_span(runs, a, b, w["entry_date"])
    a = _find(runs, "现居住址", b, exact=True)
    b = _find(runs, "\n", a + 1)
    if a < 0 or b < 0:
        raise ValueError("I could not find the place for the address in your contract paper.")
    if w.get("province"):
        _set_span(runs, a, b, w["province"])
    # ----- Khmer clause 1: dates, salary, allowance (same paragraph)
    a = _find(runs, "ខែចាប់ពីថ្ងៃទី", b)
    b = _find(runs, "នឹងទទួលបានប្រាក់ខែគោលចំនួន", a + 1)
    dg = _digits(runs, a, b)
    if a < 0 or b < 0 or len(dg) != 6:
        raise ValueError("I could not read the contract dates in the Khmer clause of your paper.")
    for r, v in zip(dg, [f"{sd:%d}", f"{sd:%m}", f"{sd:%Y}", f"{ed:%d}", f"{ed:%m}", f"{ed:%Y}"]):
        r.text = v + (NB if r.text.endswith(NB) else "")
    c2 = _find(runs, "ដុលារ", b + 1)
    _set_span(runs, b, c2, str(salary))
    a2 = _find(runs, "ប្រាក់ឧបត្ថម្ភមុខតំណែង", c2)
    a3 = _find(runs, "ចំនួន", a2 + 1)
    c3 = _find(runs, "ដុលារ", a3 + 1)
    if min(a2, a3, c3) < 0:
        raise ValueError("I could not find the allowance in the Khmer clause of your paper.")
    _set_span(runs, a3, c3, str(allowance))
    # ----- Chinese clause (一)
    cr = _para(d, "有效期从").runs
    a = _find(cr, "有效期从")
    b = _find(cr, "，将获得底薪每月为", a + 1)
    if a < 0 or b < 0:
        raise ValueError("I could not read the contract dates in the Chinese clause of your paper.")
    for r in cr[a + 1:b]:
        r.text = ""
    cr[a + 1].text = f"{NB}{NB}{sd:%Y}{NB}{NB}年{NB}{NB}{sd:%m}{NB}{NB}月{NB}{sd:%d}{NB}日 到至{NB}{NB}{ed:%Y}{NB}年{NB}{ed:%m}{NB}月{NB}{ed:%d}{NB}日"
    e1 = _find(cr, "美元", b + 1)
    _set_span(cr, b, e1, str(salary))
    f1 = _find(cr, "岗位津贴", e1)
    f2 = _find(cr, "美元", f1 + 1)
    if min(f1, f2) < 0:
        raise ValueError("I could not find the allowance in the Chinese clause of your paper.")
    _set_span(cr, f1, f2, str(allowance))
    # ----- sign date (Khmer line, then Chinese line): the day the contract starts
    for text, order in (("ធ្វើនៅ", ("d", "m", "y")), ("于", ("y", "m", "d"))):
        pr = _para(d, text).runs
        a = _find(pr, text)
        dg = _digits(pr, a, len(pr))
        if len(dg) < 3:
            raise ValueError("I could not read the sign date at the bottom of your paper.")
        val = {"d": f"{sd:%d}", "m": f"{sd:%m}", "y": f"{sd:%Y}"}
        for r, k in zip(dg[:3], order):
            r.text = val[k]
    out = io.BytesIO()
    d.save(out)
    return out.getvalue()
