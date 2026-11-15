"""Undo / Redo for saved work, safety copies, and 'open where I stopped'.

How it works (plain words):
  * Before every SAVE (Step 3 decisions, Excel save, HR changes ...) a small copy of the database is put in backup/undo.
    The Excel file is copied too when the save writes into it.
  * "Undo last save" puts the database (and the Excel) back to how they were before that save.
    "Redo" brings your change back again. Nothing is ever lost: undo itself keeps a copy for redo.
  * Copies older than KEEP_DAYS days are deleted (the newest few are always kept).
"""
import os, re, json, shutil, sqlite3, time, functools
from datetime import datetime

UNDO_DIR = os.path.join("backup", "undo")
INDEX = os.path.join(UNDO_DIR, "index.json")
STATE = os.path.join("memory", "ui_state.json")
MAX_STEPS = 25
KEEP_DAYS = 30
DEBOUNCE = 4.0                    # several writes inside one button press count as ONE undo step
_LAST = [0.0]

NICE = {
    "case_update": "Question to ask",
    "save_decisions": "Step 3 decisions", "letters_update": "letters", "outage_add": "fingerprint outage note", "outage_delete": "fingerprint outage note",
    "shift_set": "remembered shifts", "leave_add": "leave note", "leave_delete": "leave note", "letter_create": "missed-punch letter",
    "save_worker": "worker record", "save_docs": "worker documents", "save_contract": "contract", "delete_contract": "contract",
    "hp_save": "holiday-paid list", "hp_delete": "holiday-paid list", "set_setting": "setting", "save_request": "recruitment request",
    "delete_request": "recruitment request", "add_candidate": "candidate", "update_candidates": "candidates", "delete_candidate": "candidate",
    "hire_to_worker": "new worker from recruitment", "save_resign": "resignation", "mark_resigned": "resignation", "save_plan": "plan",
    "add_reminder": "reminder", "set_reminders_done": "reminder", "add_warning": "warning letter", "delete_warning": "warning letter",
    "add_factory_doc": "factory document", "delete_factory_doc": "factory document", "import_commit": "worker import", "add_doc_type": "document type",
    "times_set": "work times of a worker", "times_clear": "work times of a worker", "leave_add_photos": "photos of a leave paper", "photo_remove": "photo of a paper",
    "letter_photo_save": "photo of a letter", "resigned_set": "resigned worker", "resigned_forget": "resigned worker",
}


def _load():
    try:
        with open(INDEX, encoding="utf-8") as f:
            d = json.load(f)
        d.setdefault("undo", []); d.setdefault("redo", [])
        return d
    except Exception:
        return {"undo": [], "redo": []}


def _save(idx):
    os.makedirs(UNDO_DIR, exist_ok=True)
    tmp = INDEX + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False)
    os.replace(tmp, INDEX)


def _db_copy(src, dst):
    a = sqlite3.connect(src)
    b = sqlite3.connect(dst)
    try:
        with b:
            a.backup(b)
    finally:
        a.close(); b.close()


def _rm(entry):
    for k in ("file", "xfile"):
        p = entry.get(k)
        if p:
            try:
                os.remove(os.path.join(UNDO_DIR, p))
            except OSError:
                pass


def _capture(db_path, label, excel_path=None):
    """Copy the database (and the Excel) as they are NOW. Returns an index entry."""
    os.makedirs(UNDO_DIR, exist_ok=True)
    stamp = f"{datetime.now():%Y%m%d_%H%M%S}_{int((time.time() % 1) * 1000):03d}"
    ent = {"label": label, "at": datetime.now().isoformat(timespec="seconds"), "file": f"db_{stamp}.db"}
    _db_copy(db_path, os.path.join(UNDO_DIR, ent["file"]))
    if excel_path and os.path.exists(excel_path):
        ent["xfile"] = f"xl_{stamp}.xlsx"
        ent["xpath"] = excel_path
        shutil.copy2(excel_path, os.path.join(UNDO_DIR, ent["xfile"]))
    return ent


def snapshot(db_path, label, excel_path=None, force=False):
    """Call this just BEFORE a save. Never raises."""
    try:
        if not force and time.time() - _LAST[0] < DEBOUNCE:
            return False
        idx = _load()
        idx["undo"].append(_capture(db_path, label, excel_path))
        for e in idx["redo"]:
            _rm(e)
        idx["redo"] = []                                   # a new change ends the old redo line
        while len(idx["undo"]) > MAX_STEPS:
            _rm(idx["undo"].pop(0))
        _save(idx)
        _LAST[0] = time.time()
        return True
    except Exception:
        return False


def peek(kind="undo"):
    lst = _load()[kind]
    return lst[-1] if lst else None


def _restore(ent, db_path):
    src = os.path.join(UNDO_DIR, ent["file"])
    if not os.path.exists(src):
        raise FileNotFoundError("the saved copy is missing")
    _db_copy(src, db_path)
    if ent.get("xfile") and ent.get("xpath"):
        shutil.copy2(os.path.join(UNDO_DIR, ent["xfile"]), ent["xpath"])


def _swap(db_path, src_kind, dst_kind):
    idx = _load()
    if not idx[src_kind]:
        return None, "Nothing to " + ("undo" if src_kind == "undo" else "redo") + "."
    ent = idx[src_kind][-1]
    try:
        now = _capture(db_path, ent["label"], ent.get("xpath"))          # the state we leave: it becomes the other direction
        now["at"] = datetime.now().isoformat(timespec="seconds")
        _restore(ent, db_path)
    except Exception as e:
        return None, f"Could not do it: {type(e).__name__}: {str(e)[:150]}"
    idx[src_kind].pop()
    _rm(ent)
    idx[dst_kind].append(now)
    _save(idx)
    _LAST[0] = 0.0
    return ent, ""


def undo(db_path):
    return _swap(db_path, "undo", "redo")


def redo(db_path):
    return _swap(db_path, "redo", "undo")


def wrap(module, names, db_path):
    """Make every function in `names` take a safety copy first (once per button press)."""
    for n in names:
        fn = getattr(module, n, None)
        if fn is None or getattr(fn, "_uk_wrapped", False):
            continue

        def mk(fn=fn, n=n):
            @functools.wraps(fn)
            def inner(*a, **k):
                snapshot(db_path, NICE.get(n, n.replace("_", " ")))
                return fn(*a, **k)
            inner._uk_wrapped = True
            return inner
        setattr(module, n, mk())


def prune(days=KEEP_DAYS, keep_newest=3):
    """Delete safety copies older than `days` days. The newest few of each kind are always kept."""
    cutoff = time.time() - days * 86400
    n = 0
    for folder in ("backup", UNDO_DIR):
        if not os.path.isdir(folder):
            continue
        groups = {}
        for f in os.listdir(folder):
            p = os.path.join(folder, f)
            if not os.path.isfile(p) or f == "index.json":
                continue
            base = re.sub(r"_\d{8}_\d{6}(_\d{3})?\.[A-Za-z0-9]+$", "", f)
            groups.setdefault(base, []).append((os.path.getmtime(p), p))
        for lst in groups.values():
            lst.sort(reverse=True)
            for mt, p in lst[keep_newest:]:
                if mt < cutoff:
                    try:
                        os.remove(p); n += 1
                    except OSError:
                        pass
    if os.path.isdir(UNDO_DIR):                             # drop index entries whose files are gone
        idx = _load()
        for kind in ("undo", "redo"):
            idx[kind] = [e for e in idx[kind] if os.path.exists(os.path.join(UNDO_DIR, e["file"]))]
        _save(idx)
    return n


# ---------------------------------------------------------------- open where I stopped
def load_state():
    try:
        with open(STATE, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def save_state(d):
    try:
        if load_state() == d:
            return
        os.makedirs("memory", exist_ok=True)
        tmp = STATE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
        os.replace(tmp, STATE)
    except Exception:
        pass


# ---------------------------------------------------------------- Ctrl+Z / Ctrl+Y
HOOK_HTML = """
<script>
(function () {
  var d = window.parent.document;
  if (d.__undoHook) return;
  d.__undoHook = true;
  d.addEventListener('keydown', function (e) {                       // Alt+Enter = Checked, Alt+Left/Right = other photo
    if (!e.altKey || e.ctrlKey || e.metaKey) return;
    var sel = e.key === 'Enter' ? '[class*="st-key-ck_"] button:not([disabled])' : e.key === 'ArrowRight' ? '.st-key-pp_next2 button:not([disabled])' : e.key === 'ArrowLeft' ? '.st-key-pp_back2 button:not([disabled])' : '';
    if (!sel) return;
    var b = d.querySelector(sel);
    if (b) { e.preventDefault(); if (d.activeElement && d.activeElement.blur) d.activeElement.blur(); setTimeout(function () { b.click(); }, 60); }
  }, true);
  d.addEventListener('keydown', function (e) {
    if (!(e.ctrlKey || e.metaKey)) return;
    var k = (e.key || '').toLowerCase();
    var isUndo = (k === 'z' && !e.shiftKey), isRedo = (k === 'y' || (k === 'z' && e.shiftKey));
    if (!isUndo && !isRedo) return;
    var a = d.activeElement, tag = ((a && a.tagName) || '').toUpperCase();
    if (tag === 'INPUT' || tag === 'TEXTAREA' || (a && a.isContentEditable)) return;      // typing in a box: leave the normal undo alone
    var w = isUndo ? 'undo' : 'redo';
    var b = d.querySelector('.st-key-' + w + '_table button:not([disabled])') || d.querySelector('.st-key-' + w + '_global button:not([disabled])');
    if (b) { e.preventDefault(); b.click(); }
  }, true);
})();
(function () {
  // Keep the page where it is after a click (adding a row, updating): Streamlit sometimes shrinks the page for a moment and the browser jumps to the top.
  var d = window.parent.document;
  if (d.__scrollKeep) return;
  d.__scrollKeep = true;
  var lastY = 0, userT = 0, clickT = 0, fixing = false;
  function main() { return d.querySelector('[data-testid="stMain"]') || d.querySelector('section.main'); }
  function user() { userT = Date.now(); }
  d.addEventListener('wheel', user, true);
  d.addEventListener('touchmove', user, true);
  d.addEventListener('keydown', function (e) { if (['PageUp','PageDown','Home','End','ArrowUp','ArrowDown',' '].indexOf(e.key) >= 0) user(); }, true);
  d.addEventListener('pointerdown', function (e) {
    var m = main();
    if (!m) return;
    if (e.target === m) { user(); return; }                                    // dragging the scroll bar
    if (e.target.closest && e.target.closest('[data-testid="stSidebar"]')) { lastY = 0; clickT = 0; return; }   // moving to another page: start at the top
    if (m.contains(e.target)) clickT = Date.now();
  }, true);
  d.addEventListener('pointermove', function (e) { if (e.buttons && e.target === main()) user(); }, true);
  function restore(m, y) {
    if (fixing) return;
    fixing = true;
    var t0 = Date.now();
    (function step() {
      if (Date.now() - userT < 400 || Date.now() - t0 > 2500) { fixing = false; return; }
      if (Math.abs(m.scrollTop - y) > 4) m.scrollTop = y;
      requestAnimationFrame(step);
    })();
  }
  function attach() {
    var m = main();
    if (!m) { setTimeout(attach, 500); return; }
    m.addEventListener('scroll', function () {
      var y = m.scrollTop, now = Date.now();
      if (fixing) return;
      if (now - userT < 400) { lastY = y; return; }
      if (y < lastY - 200 && now - clickT < 15000) { restore(m, lastY); return; }
      lastY = y;
    }, { passive: true });
  }
  attach();
})();
</script>
"""
