"""
Runs a job again and again on a timer. Leave this window open.
Press Ctrl + C to stop it.

Run it:  python scheduler.py
"""
import time

import schedule

# Saturday: run the cleaning script
from clean_sales import main as job

# Sunday: put a # in front of the line above and remove the # below
# from update_dashboard import scheduled_run as job

schedule.every(1).minutes.do(job)   # in real life: schedule.every().monday.at("07:00").do(job)

print("Scheduler started. Running once now, then every 1 minute. Press Ctrl + C to stop.")
job()
while True:
    schedule.run_pending()
    time.sleep(1)
