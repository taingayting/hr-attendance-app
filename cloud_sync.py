"""
Cloud sync for the online version (Supabase Storage, free plan).

Why it exists: a free web host forgets every file when it restarts. So the website keeps its folders
(memory, web_work, backup, reports) in a PRIVATE Supabase storage bucket:
  - when the website starts it downloads the files it is missing (pull)
  - after every change it uploads the files that changed, and removes the ones you deleted (push)

If no Supabase settings are found, everything below does nothing and the website works on your laptop only.

Settings (put them in Streamlit "Secrets", or in .streamlit/secrets.toml on your laptop):
    SUPABASE_URL      = "https://xxxx.supabase.co"
    SUPABASE_KEY      = "the service_role key"      (secret! never share it, never put it on GitHub)
    SUPABASE_BUCKET   = "attendance"                (optional, default: attendance)
    ENCRYPTION_KEY    = "..."                       (made by make_secrets.py; every file is locked with it BEFORE it leaves your computer,
                                                     so Supabase only ever stores unreadable data. Keep a copy of this key somewhere
                                                     safe: without it nobody can open the data, not even you.)
"""
import os, base64, json, threading, time

import requests

SYNC_DIRS = ["memory", "web_work", "backup", "reports"]
KEEP_BACKUPS = 15                       # keep only the newest 15 Excel backups (saves space)
_lock = threading.Lock()
_synced = {}                            # relative path -> (size, mtime_ns) of what is already in the cloud
_state = {"pulled": False}


# ------------------------------------------------------------------ settings
PLACEHOLDERS = ("xxxxxxxx", "paste-", "choose-a-long", "from make_secrets", "your-key", "your key")


def _is_placeholder(v):
    """The sample file (secrets.toml.example) has fake values. If someone copies it as it is, those lines count as NOT set."""
    v = str(v).strip().lower()
    return any(p in v for p in PLACEHOLDERS)


def _secret(name, default=""):
    v = _secret_raw(name, "")
    return default if (not v or _is_placeholder(v)) else v


def _secret_raw(name, default=""):
    try:
        import streamlit as st
        v = st.secrets.get(name)
        if v:
            return str(v)
    except Exception:
        pass
    v = os.environ.get(name)
    if v:
        return v
    try:                                 # also works outside Streamlit (for upload_my_data.py)
        import tomllib
        p = os.path.join(".streamlit", "secrets.toml")
        if os.path.exists(p):
            with open(p, "rb") as f:
                v = tomllib.load(f).get(name)
            if v:
                return str(v)
    except Exception:
        pass
    return default


def config():
    url, key = _secret("SUPABASE_URL").rstrip("/"), _secret("SUPABASE_KEY")
    if not (url and key):
        return None
    return url, key, _secret("SUPABASE_BUCKET", "attendance")


def enabled():
    return config() is not None


def _fernet():
    k = _secret("ENCRYPTION_KEY")
    if not k:
        return None
    from cryptography.fernet import Fernet
    return Fernet(k.encode())


def encryption_ready():
    try:
        return _fernet() is not None
    except Exception:
        return False


# ------------------------------------------------------------------ low level
def _key(rel):                           # file names with Chinese letters are not allowed as storage keys, so we encode them
    return base64.urlsafe_b64encode(rel.encode("utf-8")).decode().rstrip("=")


def _unkey(k):
    return base64.urlsafe_b64decode(k + "=" * (-len(k) % 4)).decode("utf-8")


def _headers(key):
    return {"Authorization": f"Bearer {key}", "apikey": key}


def _ensure_bucket(url, key, bucket):
    r = requests.post(f"{url}/storage/v1/bucket", headers=_headers(key), timeout=30,
                      json={"id": bucket, "name": bucket, "public": False})
    if r.status_code not in (200, 201, 400, 409):      # 400/409 = it already exists
        r.raise_for_status()


def _list_keys(url, key, bucket):
    keys, offset = [], 0
    while True:
        r = requests.post(f"{url}/storage/v1/object/list/{bucket}", headers=_headers(key), timeout=60,
                          json={"prefix": "", "limit": 100, "offset": offset,
                                "sortBy": {"column": "name", "order": "asc"}})
        r.raise_for_status()
        items = r.json()
        keys += [i["name"] for i in items if i.get("name")]
        if len(items) < 100:
            return keys
        offset += 100


def _upload(url, key, bucket, rel, path):
    with open(path, "rb") as f:
        # the file's own path is locked inside too, so a stolen/swapped file can never pass as another file
        data = _fernet().encrypt(json.dumps({"p": rel}).encode() + b"\n" + f.read())   # locked before it leaves this computer
    h = {**_headers(key), "x-upsert": "true", "Content-Type": "application/octet-stream"}
    r = requests.post(f"{url}/storage/v1/object/{bucket}/{_key(rel)}", headers=h, data=data, timeout=180)
    r.raise_for_status()


def _download(url, key, bucket, rel, path):
    r = requests.get(f"{url}/storage/v1/object/authenticated/{bucket}/{_key(rel)}", headers=_headers(key), timeout=180)
    r.raise_for_status()
    try:
        plain = _fernet().decrypt(r.content)
    except Exception:
        raise RuntimeError("Cannot unlock a stored file. Is ENCRYPTION_KEY the same key that was used when the data was sent?")
    head, _, body = plain.partition(b"\n")
    try:
        ok = json.loads(head).get("p") == rel
    except Exception:
        ok = False
    if not ok:
        raise RuntimeError(f"A stored file does not match its name ({rel}). It may have been swapped. I stopped to protect your data.")
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "wb") as f:
        f.write(body)


def _delete(url, key, bucket, rels):
    for i in range(0, len(rels), 50):
        r = requests.delete(f"{url}/storage/v1/object/{bucket}", headers=_headers(key), timeout=60,
                            json={"prefixes": [_key(x) for x in rels[i:i + 50]]})
        r.raise_for_status()


def _skip(rel):
    base = os.path.basename(rel)
    return base.startswith("_incoming_") or base.startswith("~$") or base.endswith(".tmp") or base.endswith("-journal")


def _local_files():
    out = {}
    for d in SYNC_DIRS:
        for root, _, files in os.walk(d):
            for f in files:
                rel = os.path.join(root, f).replace(os.sep, "/")
                if _skip(rel):
                    continue
                st = os.stat(rel)
                out[rel] = (st.st_size, st.st_mtime_ns)
    return out


def _prune_photos(days=7):
    """Working copies of paper photos older than 7 days are removed (every checked day keeps its photos in web_work/archive)."""
    now = time.time()
    for folder in (os.path.join("web_work", "photos"), os.path.join("web_work", "photo_work")):
        if not os.path.isdir(folder):
            continue
        for f in os.listdir(folder):
            p = os.path.join(folder, f)
            try:
                if os.path.isfile(p) and now - os.path.getmtime(p) > days * 86400:
                    os.remove(p)
            except OSError:
                pass


def _prune_backups():
    """Keep only the newest KEEP_BACKUPS Excel backups on disk (the database backups are small, they stay)."""
    files = sorted((os.path.join("backup", f) for f in os.listdir("backup")
                    if f.lower().endswith((".xlsx", ".xlsm"))), key=os.path.getmtime) if os.path.isdir("backup") else []
    for old in files[:-KEEP_BACKUPS]:
        try:
            os.remove(old)
        except OSError:
            pass


# ------------------------------------------------------------------ public
def pull():
    """Once per start: download the files that are in the cloud but not on this machine. Returns how many files came down."""
    cfg = config()
    if not cfg or _state["pulled"]:
        return 0
    if not encryption_ready():
        raise RuntimeError("ENCRYPTION_KEY is missing or not valid")
    url, key, bucket = cfg
    with _lock:
        if _state["pulled"]:
            return 0
        _ensure_bucket(url, key, bucket)
        n = 0
        for k in _list_keys(url, key, bucket):
            try:
                rel = _unkey(k)
            except Exception:
                continue
            if not rel.split("/")[0] in SYNC_DIRS or _skip(rel):
                continue
            if not os.path.exists(rel):
                _download(url, key, bucket, rel, rel)
                n += 1
            if os.path.exists(rel):
                st = os.stat(rel)
                _synced[rel] = (st.st_size, st.st_mtime_ns)
        _state["pulled"] = True
        return n


def push():
    """Upload what changed since the last push, and delete from the cloud what was deleted here. Returns (uploaded, deleted)."""
    cfg = config()
    if not cfg or not _state["pulled"]:
        return 0, 0
    url, key, bucket = cfg
    with _lock:
        _prune_backups()
        _prune_photos()
        now = _local_files()
        up = [r for r, sig in now.items() if _synced.get(r) != sig]
        gone = [r for r in list(_synced) if r not in now]
        for rel in up:
            _upload(url, key, bucket, rel, rel)
            _synced[rel] = now[rel]
        if gone:
            _delete(url, key, bucket, gone)
            for rel in gone:
                _synced.pop(rel, None)
        return len(up), len(gone)


def push_all_first_time():
    """For upload_my_data.py: send everything on this laptop to the cloud (without downloading first)."""
    cfg = config()
    if not cfg:
        raise SystemExit("No Supabase settings found. Create .streamlit/secrets.toml first (see the setup guide).")
    if not encryption_ready():
        raise SystemExit("ENCRYPTION_KEY is missing in .streamlit/secrets.toml. Run  py make_secrets.py  and add the key.")
    url, key, bucket = cfg
    _ensure_bucket(url, key, bucket)
    _state["pulled"] = True
    return push()