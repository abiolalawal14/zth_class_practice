"""
FreshMart weekly sales cleaner.

Reads every branch file in data/raw, fixes the 6 problems we found in class,
and saves a dated Excel report to the output folder.

Run it:  python clean_sales.py
"""
import logging
from datetime import date
from pathlib import Path

import pandas as pd

# Folders - worked out from where this file lives, so the script works
# no matter where it is run from (important for the scheduler).
FOLDER = Path(__file__).resolve().parent
RAW_FOLDER = FOLDER / "data" / "raw"
OUTPUT_FOLDER = FOLDER / "output"
OUTPUT_FOLDER.mkdir(exist_ok=True)

# Problem 1: each branch names its columns differently
COLUMN_NAMES = {
    "date_sold": "date",
    "qty": "quantity",
    "price": "unit_price",
    "payment_method": "payment",
}
# Problem 2: typing mistakes in branch names
BRANCH_TYPOS = {"Ikj": "Ikeja"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(OUTPUT_FOLDER / "automation_log.txt", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("freshmart")


def load_all_files(folder=RAW_FOLDER):
    """Read every Excel file in the folder and stack them into one table."""
    frames = []
    for path in sorted(folder.glob("*.xlsx")):
        if path.name.startswith("~$"):  # temporary file Excel makes while a file is open
            continue
        df = pd.read_excel(path)
        df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
        df = df.rename(columns=COLUMN_NAMES)
        frames.append(df)
        log.info(f"Read {path.name}: {len(df)} rows")
    if not frames:
        raise FileNotFoundError(f"No Excel files found in {folder}")
    return pd.concat(frames, ignore_index=True)


def clean(sales):
    """Fix the messy data. Returns the clean table."""
    sales = sales.copy()

    # Problem 2: branch names ("ikeja ", "IKEJA", "Ikj", "WUSE II" -> one spelling)
    sales["branch"] = sales["branch"].str.strip().str.title()
    sales["branch"] = sales["branch"].str.split().str[0]
    sales["branch"] = sales["branch"].replace(BRANCH_TYPOS)

    # Problem 3: dates written in different ways
    sales["date"] = pd.to_datetime(sales["date"], format="mixed", dayfirst=True)

    # Problem 4: prices stored as text ("₦9,500") - keep only the digits
    sales["unit_price"] = sales["unit_price"].astype(str).str.replace(r"[^0-9.]", "", regex=True)
    sales["unit_price"] = pd.to_numeric(sales["unit_price"])

    # Problem 5: the same sale typed twice
    duplicates = sales.duplicated().sum()
    sales = sales.drop_duplicates()
    log.info(f"Removed {duplicates} duplicate rows")

    # Problem 6: missing quantity - we can't guess a sale, so set it aside for checking
    missing = sales[sales["quantity"].isna()]
    if len(missing) > 0:
        missing.to_excel(OUTPUT_FOLDER / "rows_to_check.xlsx", index=False)
        log.info(f"{len(missing)} rows have no quantity - saved to output/rows_to_check.xlsx")
    sales = sales.dropna(subset=["quantity"])
    sales["quantity"] = sales["quantity"].astype(int)

    sales["revenue"] = sales["quantity"] * sales["unit_price"]
    sales = sales.sort_values(["date", "branch"]).reset_index(drop=True)
    sales.attrs["rows_to_check"] = len(missing)  # remembered so the n8n message can mention it
    return sales


def summarise(sales):
    """Build the summary tables managers care about."""
    by_branch = (sales.groupby("branch")["revenue"].sum()
                 .sort_values(ascending=False).reset_index())
    by_product = (sales.groupby("product")
                  .agg(units_sold=("quantity", "sum"), revenue=("revenue", "sum"))
                  .sort_values("revenue", ascending=False).reset_index())
    by_day = sales.groupby("date")["revenue"].sum().reset_index()
    by_payment = sales.groupby("payment")["revenue"].sum().reset_index()
    return {"by_branch": by_branch, "by_product": by_product,
            "by_day": by_day, "by_payment": by_payment}


def save_report(sales, summaries):
    """Save a dated report, plus a 'latest' copy that Power BI always reads."""
    today = date.today().isoformat()
    dated_file = OUTPUT_FOLDER / f"freshmart_report_{today}.xlsx"
    latest_file = OUTPUT_FOLDER / "freshmart_clean_latest.xlsx"
    for path in (dated_file, latest_file):
        with pd.ExcelWriter(path) as writer:
            sales.to_excel(writer, sheet_name="clean_data", index=False)
            for name, table in summaries.items():
                table.to_excel(writer, sheet_name=name, index=False)
    log.info(f"Saved {dated_file.name} and {latest_file.name}")
    return dated_file


def main():
    log.info("---- FreshMart cleaning started ----")
    sales = clean(load_all_files())
    summaries = summarise(sales)
    save_report(sales, summaries)
    log.info(f"Done: {len(sales)} clean rows, total revenue ₦{sales['revenue'].sum():,.0f}")
    return sales, summaries


if __name__ == "__main__":
    main()
