# HR Attendance App

**An AI helper that checks a factory's daily attendance, so workers are paid correctly and HR finishes on time.**

👉 **Try the real app:** https://hr-attendance-app-lxvsk26tyd4fqtfdmm28jp.streamlit.app/
⏱ The free server sleeps when nobody uses it. The first visit can take about **1 minute** to wake up. It is not broken.

Made by [Taing Ayting](https://taingayting.github.io/portfolio/). I use the private version of this app in my HR job every day.

---

## Why I built it

When I started in HR at a factory in Cambodia, I knew nothing about HR, and I was alone with a lot of work.

Every day, two records must agree:

- the **paper sheet** that supervisors fill in by hand (who worked, and how many overtime hours), and
- the **fingerprint machine**, which records when each worker came in and went out.

Someone has to compare them, worker by worker, and copy the result into a big monthly Excel file. Workers' **pay depends on it**. Doing it by hand was slow, tiring and easy to get wrong. I needed better management to finish on time, so I built this app.

## Why it matters

- **Pay must be right.** One wrong cell can mean a worker is paid too little. The app checks every worker, every day.
- **People stay in control.** AI helps read the paper, but a person makes every final decision. Nothing is saved until HR types **YES**.
- **Small factories cannot pay for big HR software.** This app runs on free tools and an ordinary laptop.
- **It is built from real work.** I wrote the rules from my own job and changed them when real data proved me wrong.

## What it does (the core)

### Attendance (the daily work)

1. **Reads the fingerprint export** (CSV or Excel) and the monthly attendance Excel.
2. **Reads the paper sheet.** Type it, paste it, or upload a photo. An AI model (Claude or Google Gemini) reads the handwriting. There is also a local reader (RapidOCR) that works without internet.
3. **Compares paper and fingerprint for every worker.**
   - 🟢 Green: paper and fingerprint agree.
   - 🟡 Yellow: they are different. The app shows why.
   - 🔴 Red / "?": the app is not sure, so it asks a person. It never guesses.
4. **Shows every cell it wants to write**, and waits until HR types **YES**.
5. **Makes a backup, saves, then reads every cell back** to make sure the file is right. One click can undo the save.

Other attendance pages: one worker's full month, history, leave and missed-punch letters, a list of questions to ask workers, the monthly report, a scorecard, the worker list, a page for Chinese workers, and notes.

### HR (the people side)

- **Today:** what needs attention now (contracts ending, reminders, new workers).
- **Workers:** each worker's profile, documents, contract and history.
- **Contracts:** makes contract documents (Word files) from a template.
- **Recruitment:** job requests, candidates and hiring steps.
- **Resigned, Reminders, Reports and Factory documents.**

## One hard decision I made

**The app must never save on its own.** It would be faster to save automatically, but pay depends on these numbers, and AI can misread handwriting. So the app shows every change first, a person types YES, and the app checks the file again after saving. Speed matters, but trust matters more.

## What real use taught me

The written rule said the night shift ends at 03:30. When I checked one week of real fingerprint data, **232 of 441 night shifts ended near 06:00**. So I changed the rules to follow what really happens, not only what is written.

## How to use the live app

1. Open the link above. You start on **Daily work**, with **14 sample workers** on Monday 5 October 2026.
2. The fingerprint file and the paper list are already loaded. Press the button to **check**.
3. Look at the colours. Open a yellow or red worker to see why.
4. Go to the save step. Read the list of cells, then type **YES** to save.
5. Try the other pages in the sidebar: **Attendance** and **HR**.

This public copy is safe to play with. Everything resets after 30 minutes without visitors.

## Tools I used

| Part | Tool |
| :--- | :--- |
| Language | Python 3.12 |
| Web app | Streamlit |
| Data | pandas, SQLite |
| Excel files | openpyxl (reads and writes the real factory format) |
| Word contracts | python-docx (fills your own contract template) |
| Reading paper photos | Claude or Google Gemini (vision), RapidOCR + OpenCV (offline) |
| Security (private version) | Password + 6-digit phone code (pyotp), Fernet encryption (cryptography), lockout and login history |
| Building | AI coding tools, mainly Claude, wrote much of the code with me. I found the problem, wrote the rules, tested with real files and decided what is right. |

The private version also syncs encrypted files to cloud storage. See [SECURITY.md](SECURITY.md) for how it protects worker data.

## Size

- **14,000+ lines** of Python
- **9,000+ saved changes** logged in my real use

## Privacy

My real workers' data is private and is **not** in this repository. This public copy uses **14 invented workers**. Their names are "Sample Worker 01" and so on, and their IDs start with `DMO-`. The sample data is created fresh by `demo_seed.py` every time the app starts.

## Run it on your computer

You need Python 3.12 or newer.

```
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Then open http://localhost:8501.

## Files

| File | What it does |
| :--- | :--- |
| `streamlit_app.py` | Starts the public demo: creates the sample data, then runs the app |
| `app.py` | The main app: pages, the daily check, the save step |
| `attendance_agent.py` | The rules that compare paper and fingerprint |
| `paper_reader.py`, `local_reader.py` | Read paper sheets from photos (AI or offline) |
| `attendance_db.py`, `hr_db.py` | Databases for attendance and HR records |
| `hr_pages.py`, `person_page.py`, `today_panel.py` | HR pages |
| `contract_docx.py` | Makes contract documents |
| `monthly_report.py`, `report_xlsx.py`, `scorecard.py` | Reports |
| `undo_kit.py` | Backups and undo |
| `demo_seed.py` | The 14 sample workers |

