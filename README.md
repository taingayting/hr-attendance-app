# HR Attendance App

I started in HR at a factory in Cambodia knowing nothing, and I had to finish a lot of work on time. Every day, workers fill in paper attendance sheets and scan their fingerprints. Someone must compare the two, worker by worker, and workers' pay depends on the result. So I built this app to help me do it.

**Try it:** [live app link, added after deploy]

This public copy runs on **14 sample workers**. My real workers' data is private and is not in this repository. Worker IDs here start with `DMO-`.

## What it does

1. Reads the fingerprint export and the monthly attendance Excel.
2. Takes the paper list (typed in, or read from a photo by an AI model).
3. Compares paper and fingerprint for every worker. Anything unclear becomes "?" and goes to a person.
4. Shows every cell it wants to write, and waits until HR types **YES**.
5. Makes a backup, saves, then reads every cell back to confirm.

It also has an HR side: worker records, contracts, recruitment, reminders, letters and reports.

## How I built it

I found the problem in my own job, wrote the rules, and tested them on real files. AI coding tools, mainly Claude, wrote much of the code with me. Real data changed my rules: in one week, 232 of 441 night shifts ended near 06:00, not at 03:30 as the written rule said.

## Run it on your computer

```
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Python 3.12 or newer.
