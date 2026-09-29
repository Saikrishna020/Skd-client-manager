"""
Health Open Cases report - builds two outputs from a "Health open cases"
export:

  1. Client-wise breakdown: one row per Client, with Total cases, a
     count per Sub Product found in the file, and:
       - FO Closed: count where Status == "FO Completed"
       - Pending: count where Status == "Pending"

  2. Manager Pending Closure: one row per Manager, restricted to cases
     where Status == "FO Completed" (FO has finished fieldwork, case is
     waiting on the manager to close it), with:
       - Total: count of such cases for that manager
       - More than 5 days: of those, SKD TAT-D > 5
       - Cashless/Spot intimation: of those, Sub Product is "Cashless"
         or "Spot Intimation"
       - More than 1 day after closure by CAT: of those, CAT Completed
         Date is 2 or more days before today (i.e. more than 1 day ago)
       - % more than 1 day after closure by CAT: that count / Total * 100
       - % More than 5 days: that count / Total * 100

"Today" is always the actual date the report is run, not a date found
in the file.

Usage:
    python health_open_cases_report.py "Health open cases 16 Sep 26.xlsx" -o "Health_Open_Cases_Report.xlsx"

If -o is omitted, output is named "<input>_Open_Cases_Report.xlsx" and
contains both outputs as separate sheets.
"""

import argparse
import math
import sys
from datetime import datetime

import pandas as pd


def round_half_up(value: float) -> int:
    """Standard rounding (0.5 always rounds up), unlike Python's built-in
    round() which rounds 0.5 to the nearest even number."""
    return math.floor(value + 0.5)

CLIENT_COL = "Client"
SUB_PRODUCT_COL = "Sub Product"
MANAGER_COL = "Manager"
STATUS_COL = "Status"
SKD_TAT_D_COL = "SKD TAT-D"
CAT_COMPLETED_DATE_COL = "CAT Completed Date"

FO_COMPLETED_STATUS = "FO Completed"
PENDING_STATUS = "Pending"
SKD_TAT_D_THRESHOLD = 5
CAT_CLOSURE_DAYS_THRESHOLD = 1  # "more than 1 day" -> elapsed > 1

CASHLESS_SPOT_SUB_PRODUCTS = {"cashless", "spot intimation"}


def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Sheet1", header=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    return df


def build_client_report(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(subset=[CLIENT_COL])
    sub_products = sorted(df[SUB_PRODUCT_COL].dropna().unique())

    clients = sorted(df[CLIENT_COL].dropna().unique())
    rows = []
    for client in clients:
        sub = df[df[CLIENT_COL] == client]
        row = {"Client Name": client, "Total Cases": len(sub)}
        for sp in sub_products:
            row[sp] = int((sub[SUB_PRODUCT_COL] == sp).sum())
        row["FO Closed"] = int((sub[STATUS_COL] == FO_COMPLETED_STATUS).sum())
        row["Pending"] = int((sub[STATUS_COL] == PENDING_STATUS).sum())
        rows.append(row)

    report = pd.DataFrame(rows)

    total_row = {"Client Name": "Grand Total", "Total Cases": report["Total Cases"].sum()}
    for sp in sub_products:
        total_row[sp] = report[sp].sum()
    total_row["FO Closed"] = report["FO Closed"].sum()
    total_row["Pending"] = report["Pending"].sum()
    report = pd.concat([report, pd.DataFrame([total_row])], ignore_index=True)

    return report


def build_manager_pending_closure_report(df: pd.DataFrame, today: datetime = None) -> pd.DataFrame:
    if today is None:
        today = datetime.now()

    df = df.dropna(subset=[MANAGER_COL])
    fo_completed = df[df[STATUS_COL] == FO_COMPLETED_STATUS].copy()

    cat_completed = pd.to_datetime(
        fo_completed[CAT_COMPLETED_DATE_COL], dayfirst=True, errors="coerce"
    )
    elapsed_days = (pd.Timestamp(today.date()) - cat_completed).dt.days
    fo_completed["_more_than_5_days"] = fo_completed[SKD_TAT_D_COL] > SKD_TAT_D_THRESHOLD
    fo_completed["_cashless_spot"] = (
        fo_completed[SUB_PRODUCT_COL].str.strip().str.lower().isin(CASHLESS_SPOT_SUB_PRODUCTS)
    )
    fo_completed["_more_than_1_day_after_cat"] = elapsed_days > CAT_CLOSURE_DAYS_THRESHOLD

    managers = sorted(fo_completed[MANAGER_COL].dropna().unique())
    rows = []
    for manager in managers:
        sub = fo_completed[fo_completed[MANAGER_COL] == manager]
        total = len(sub)
        more_than_5_days = int(sub["_more_than_5_days"].sum())
        more_than_1_day_after_cat = int(sub["_more_than_1_day_after_cat"].sum())
        rows.append(
            {
                "Manager Name": manager,
                "Total": total,
                "More than 5 days": more_than_5_days,
                "Cashless/Spot intimation": int(sub["_cashless_spot"].sum()),
                "More than 1 day after closure by CAT": more_than_1_day_after_cat,
                "% more than 1 day after closure by CAT": round_half_up(more_than_1_day_after_cat / total * 100) if total else 0,
                "% More than 5 days": round_half_up(more_than_5_days / total * 100) if total else 0,
            }
        )

    report = pd.DataFrame(rows)

    total_total = report["Total"].sum()
    total_more_than_5_days = report["More than 5 days"].sum()
    total_more_than_1_day_after_cat = report["More than 1 day after closure by CAT"].sum()
    total_row = {
        "Manager Name": "Grand Total",
        "Total": total_total,
        "More than 5 days": total_more_than_5_days,
        "Cashless/Spot intimation": report["Cashless/Spot intimation"].sum(),
        "More than 1 day after closure by CAT": total_more_than_1_day_after_cat,
        "% more than 1 day after closure by CAT": round_half_up(total_more_than_1_day_after_cat / total_total * 100) if total_total else 0,
        "% More than 5 days": round_half_up(total_more_than_5_days / total_total * 100) if total_total else 0,
    }
    report = pd.concat([report, pd.DataFrame([total_row])], ignore_index=True)

    return report


def main():
    parser = argparse.ArgumentParser(description="Build Health Open Cases reports.")
    parser.add_argument("input", help="Path to the source .xlsx file")
    parser.add_argument("-o", "--output", help="Path to write the report .xlsx file")
    args = parser.parse_args()

    output = args.output
    if not output:
        if args.input.lower().endswith(".xlsx"):
            output = args.input[:-5] + "_Open_Cases_Report.xlsx"
        else:
            output = args.input + "_Open_Cases_Report.xlsx"

    df = load_data(args.input)
    client_report = build_client_report(df)
    manager_report = build_manager_pending_closure_report(df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        client_report.to_excel(writer, sheet_name="Client Wise", index=False)
        manager_report.to_excel(writer, sheet_name="Manager Pending Closure", index=False)

    print(f"Report written to: {output}")
    print(f"Clients: {len(client_report) - 1}, Total cases: {client_report.iloc[-1]['Total Cases']}")
    print(f"Managers: {len(manager_report) - 1}, Total pending: {manager_report.iloc[-1]['Total']}")


if __name__ == "__main__":
    sys.exit(main())
