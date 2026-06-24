# AUTOMATED REPORT — Runs on schedule, updates Excel
# Retail Chain Sales Analytics

import schedule
import time
import subprocess
import os
from datetime import datetime

BASE_DIR = r"C:\Users\KIIT\OneDrive\Desktop\data_analytic\Sales Data Analysis and Reporting for a Retail Chain"
LOG_FILE = os.path.join(BASE_DIR, "reports", "auto_report_log.txt")

def run_report():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] Running automated report...")

    # Step 1: Re-run phase 4 to regenerate Excel
    result = subprocess.run(
        ["python",
         os.path.join(BASE_DIR, "src", "phase4_excel.py")],
        capture_output=True, text=True
    )

    # Step 2: Log the result
    with open(LOG_FILE, "a") as f:
        f.write(f"\n{'='*40}\n")
        f.write(f"Run Time : {now}\n")
        if result.returncode == 0:
            f.write(f"Status   : SUCCESS\n")
            f.write(f"Output   : {result.stdout[-200:]}\n")
            print(f"[{now}] ✅ Report updated successfully")
        else:
            f.write(f"Status   : FAILED\n")
            f.write(f"Error    : {result.stderr[-200:]}\n")
            print(f"[{now}] ❌ Report failed — check log")

# ── Schedule Options (uncomment what you want) ────────────
# Runs every day at 8:00 AM
schedule.every().day.at("08:00").do(run_report)

# Runs every Monday at 9:00 AM
# schedule.every().monday.at("09:00").do(run_report)

# Runs every hour
# schedule.every().hour.do(run_report)

# ── Run once immediately on start ─────────────────────────
print("=" * 45)
print("  AUTOMATED REPORT SCHEDULER STARTED")
print("=" * 45)
print("  Schedule : Every day at 08:00 AM")
print("  Log file : reports/auto_report_log.txt")
print("  Press Ctrl+C to stop")
print("=" * 45)

run_report()   # run immediately on start

# ── Keep running ──────────────────────────────────────────
while True:
    schedule.run_pending()
    time.sleep(60)