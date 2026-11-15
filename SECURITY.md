# Security of the Attendance helper (online version)

Honest summary: nothing connected to the internet is 100% safe. This design makes an attack hard, makes a stolen copy useless,
and makes an attempt visible. Review done: threat model, code scan (bandit: no high/medium issues), package check (pip-audit: no known vulnerabilities),
and the OWASP Top 10 (2025) list.

## Layers (an attacker must beat ALL of them)
1. **Streamlit private app** - only the email you invite can even open the page (Google sign-in or emailed link). [Setup step C5]
2. **Password** (12+ characters) AND **6-digit phone code** (changes every 30 seconds, each code works once).
3. **Lockout**: 5 wrong attempts = locked 10 minutes. **Auto logout** after 30 minutes without a click.
4. **Login history**: every success/failure is logged and shown at the next login ("wrong attempts since your last login").
5. **Encryption**: every file is locked on the server BEFORE it is sent to Supabase (Fernet = AES + tamper check). Supabase stores unreadable data.
   Each file also carries its own path inside the lock, so files cannot be swapped.
6. **Private storage + private code**: private Supabase bucket, private GitHub repo, .gitignore blocks data and secrets, secrets only in Streamlit Secrets.
7. **Safe inputs**: uploads are checked by CONTENT, file names are cleaned, downloaded CSVs cannot run Excel formulas, errors show no technical details.
8. **Pinned packages**: exact tested versions (supply-chain safety). Turn on GitHub "Dependabot alerts" for the repo (Settings -> Code security).

## OWASP Top 10 (2025) check
- A01 Broken access control: private app + password + 2-step code; single role, nothing public.
- A02 Security misconfiguration: hardened config (XSRF on, errors hidden, toolbar hidden); online mode refuses to start if a secret is missing; laptop mode listens on localhost only.
- A03 Supply-chain failures: pinned versions, pip-audit clean. Keep them updated every few months.
- A04 Cryptographic failures: encryption before upload; key only in secrets; TLS to Supabase (HTTPS).
- A05 Injection: parameterised SQL; CSV formula protection; no HTML from user data.
- A06 Insecure design: least data online, encrypted, backups, one-person use.
- A07 Authentication failures: password + TOTP, lockout, single-use codes, session timeout.
- A08 Integrity failures: Fernet tamper check + path binding. Not covered: an attacker with the Supabase key could put back an OLDER valid copy of a file. Keep your laptop backups.
- A09 Logging/alerting: security_log + login notice. No email alerts (not built).
- A10 Exceptional conditions: storage errors stop the app safely (it never runs on half-loaded data).

## Photos (paper reader) - what leaves your system
- Only when you press "Read", each photo is sent to the company of your key (Google for an AIza key, Anthropic for an sk-ant key) over an encrypted connection (HTTPS). Nothing is sent before you tick the agreement box. Every send is counted in memory/security_log.csv.
- This is the one place where worker data (IDs, names on the paper) leaves your own system. Ask your company first. **Google's FREE plan:** Google's terms for free (unpaid) Gemini API use say Google may use the content to improve its products and humans may review it. Read the current terms at ai.google.dev/gemini-api/terms. A paid plan is treated differently. **Anthropic:** read the current API privacy terms on anthropic.com for retention and training.
- The key (sent in a request header, never in a web address) lives only in the Streamlit Secrets (or your laptop's secrets.toml, blocked from GitHub). Set a monthly spending limit in the Anthropic Console so a stolen key cannot run up a bill, and delete the key there if you ever suspect it leaked.
- Photos are re-saved as plain JPEG before anything else (this removes hidden data and rejects fake pictures); very large "picture bombs" are refused.
- What Claude answers is checked by the program: IDs must look like DMO-12345, hours below 8, OT only the paper columns. Anything else becomes "?". Text written on the paper cannot give the program orders.
- Claude can still misread handwriting. That is why every photo is shown next to its table and you must press "Checked" on each photo before it can be used, and why the fingerprint check comes after.
- Saved photos are encrypted like all other files before they go to the online storage. They use space: about 5 MB per day. The free storage is 1000 MB, so look at the number in History -> Backup once a month.

## What can still go wrong (and what to do)
- **You lose ENCRYPTION_KEY** -> online data is unreadable forever. Keep it on paper. Keep backups on your laptop (the .xlsx you download).
- **Your phone or password is stolen** -> change APP_PASSWORD and TOTP_SECRET in Streamlit Secrets and re-run make_secrets.py.
- **Phishing**: only type the password on YOUR app address. Nobody (including me) should ask for it.
- **Someone with your Supabase SECRET key** can delete files (not read them). Keep it only in Streamlit Secrets. If leaked: Supabase -> rotate the key.
- **Lockout can be used against you**: an attacker can lock you out for 10 minutes by guessing. Wait, or restart the app.
- **Smart mode (AI key)** sends worker data to the AI company. Leave it empty online.
- **Streamlit/Supabase staff** run the servers. They only see unreadable files, but file names and dates are visible.
- **Law/company rules**: ask your company whether worker data may be stored online and where (Singapore region).
- Not built: email alerts, IP allow-listing, per-user accounts. Tell me if you need them.
