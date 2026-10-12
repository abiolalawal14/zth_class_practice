"""
FreshMart dashboard updater.

1. Cleans all branch files (using clean_sales.py)
2. Saves the Excel report that Power BI reads
3. Sends the clean data and summaries to the Google Sheet
4. Tells n8n how it went, so n8n can email / Telegram you

Run it:  python update_dashboard.py
Needs:   credentials.json in this folder, and the Google Sheet shared with
         the service account email (see the setup guide).
         For n8n alerts: paste your n8n webhook URL into n8n_webhook.txt
"""
import json
from datetime import datetime

import gspread
import requests

from clean_sales import FOLDER, OUTPUT_FOLDER, clean, load_all_files, log, save_report, summarise

SHEET_NAME = "FreshMartDashboard"
CREDENTIALS = FOLDER / "credentials.json"
WEBHOOK_FILE = FOLDER / "n8n_webhook.txt"
LAST_RUN_FILE = OUTPUT_FOLDER / "last_run.json"

TABS = {
    "by_branch": "By Branch",
    "by_product": "By Product",
    "by_day": "By Day",
    "by_payment": "By Payment",
}


def get_tab(sheet, name):
    """Open a tab by name, or create it if it doesn't exist yet."""
    try:
        return sheet.worksheet(name)
    except gspread.WorksheetNotFound:
        return sheet.add_worksheet(title=name, rows=1000, cols=20)


def send_table(sheet, tab_name, table):
    """Replace everything on a tab with this table."""
    rows = [table.columns.tolist()] + table.astype(str).values.tolist()
    tab = get_tab(sheet, tab_name)
    tab.clear()
    tab.update(values=rows, range_name="A1", value_input_option="USER_ENTERED")
    log.info(f"Sent {len(table)} rows to the '{tab_name}' tab")


def notify_n8n(message):
    """Send a message to the n8n webhook. Skipped if no URL has been pasted in yet."""
    url = WEBHOOK_FILE.read_text(encoding="utf-8").strip() if WEBHOOK_FILE.exists() else ""
    if not url.startswith("http"):
        log.info("No n8n webhook URL in n8n_webhook.txt - skipping the alert")
        return
    try:
        response = requests.post(url, json=message, timeout=15)
        log.info(f"Told n8n: {message['status']} (n8n replied {response.status_code})")
    except requests.RequestException as error:
        log.info(f"Could not reach n8n: {error}")


def data_changed(summary):
    """True if the numbers are different from the last run (so we don't spam alerts)."""
    now = {"clean_rows": summary["clean_rows"], "total_revenue": summary["total_revenue"]}
    before = json.loads(LAST_RUN_FILE.read_text()) if LAST_RUN_FILE.exists() else None
    LAST_RUN_FILE.write_text(json.dumps(now))
    return now != before


def update():
    """Clean, save, send to Google Sheets. Returns a short summary."""
    sales = clean(load_all_files())
    summaries = summarise(sales)
    save_report(sales, summaries)

    gc = gspread.service_account(filename=r"C:\Users\HP\Documents\FreshMart_Automation_Class\credentials.json")
    sheet = gc.open(SHEET_NAME)
    send_table(sheet, "Clean Data", sales)
    for key, tab_name in TABS.items():
        send_table(sheet, tab_name, summaries[key])

    stamp = datetime.now().strftime("%d %b %Y, %I:%M %p")
    get_tab(sheet, "By Branch").update(values=[["Last updated", stamp]], range_name="D1")
    log.info(f"Google Sheet updated: {sheet.url}")

    by_branch = summaries["by_branch"]
    total = int(sales["revenue"].sum())
    return {
        "status": "success",
        "time": stamp,
        "branches": len(by_branch),
        "branch_list": ", ".join(sorted(by_branch["branch"])),
        "top_branch": by_branch.iloc[0]["branch"],
        "clean_rows": len(sales),
        "rows_to_check": sales.attrs.get("rows_to_check", 0),
        "total_revenue": total,
        "total_revenue_text": f"₦{total:,}",
        "sheet_url": sheet.url,
    }


def main(always_notify=True):
    """Run the update and tell n8n. The scheduler uses always_notify=False:
    then n8n only hears from us when the numbers changed, or when something broke."""
    log.info("---- FreshMart dashboard update started ----")
    try:
        summary = update()
    except Exception as error:
        log.exception("Update FAILED")
        notify_n8n({"status": "failed",
                    "time": datetime.now().strftime("%d %b %Y, %I:%M %p"),
                    "error": f"{type(error).__name__}: {error}"})
        return None
    if data_changed(summary) or always_notify:
        notify_n8n(summary)
    else:
        log.info("Nothing changed since last run - no alert sent")
    return summary


def scheduled_run():
    main(always_notify=False)


if __name__ == "__main__":
    main()
