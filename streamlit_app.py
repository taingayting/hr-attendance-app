"""Public demo entry point: runs the real app on sample workers.

Everyone who opens the demo shares one copy of the sample data. It goes back to a fresh day
after 30 minutes without any click, so the next visitor can try the whole flow again.
"""
import os
import runpy
import shutil
import tempfile
import time

import demo_seed

DATA = os.path.join(tempfile.gettempdir(), "hr-attendance-demo-data")
STAMP = os.path.join(DATA, ".last_click")
RESET_AFTER = 30 * 60

here = os.path.dirname(os.path.abspath(__file__))
os.chdir(here)
idle = time.time() - os.path.getmtime(STAMP) if os.path.exists(STAMP) else None
if idle is None or idle > RESET_AFTER:
    shutil.rmtree(DATA, ignore_errors=True)
    os.makedirs(DATA, exist_ok=True)
    demo_seed.seed(DATA)
with open(STAMP, "w") as f:
    f.write(str(time.time()))

os.environ["ATT_DATA"] = DATA
os.environ["HR_DEMO"] = "1"
runpy.run_path(os.path.join(here, "app.py"), run_name="__main__")
