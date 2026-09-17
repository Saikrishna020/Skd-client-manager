"""
Manager-wise TAT compliance report.

Reads the "Sheet1" data from the SKD Corinsoft "Health Managers closed"
export (header on row 1, data from row 2), groups cases by Manager and
Sub Product, and reports, for each Manager:
    - Total cases handled
    - Total cases completed within TAT
    - Total cases Discrepant
    - Total cases with Manager closed days >= 2
    - Per category: case count, within-TAT count, Discrepant count, and
      Manager closed days >= 2 count
      (Cashless, Cashless Full case, Benefit, FULL INVESTICATION, MBV,
       PRE CLAIM VERIFICATION, Re Investigation, Reimbursement,
       Reimbursement-Half case, TP)

"Within TAT" = TAT (days) <= the standard TAT for that sub product
(defined in STANDARD_TAT below).

"Discrepant" = Final Conclusion column equals "Discrepant".

"Manager closed days >= 2" = Manager closed days column value is 2 or
more (i.e. the manager took 2+ days to close it).

Usage:
    python manager_tat_report.py "Health Managers closed  01 Apr to 31 Aug 26.xlsx" -o "Manager_TAT_Report.xlsx"

If -o is omitted, the output file is named "<input>_Manager_TAT_Report.xlsx".
"""

import argparse
import sys

import pandas as pd

# Standard TAT (in days) allowed per Sub Product.
STANDARD_TAT = {
    "Cashless": 1,
    "Cashless Full case": 5,
    "Benefit": 5,
    "FULL INVESTICATION": 7,
    "MBV": 3,
    "PRE CLAIM VERIFICATION": 3,
    "Re Investigation": 3,
    "Reimbursement": 5,
    "Reimbursement-Half case": 5,
    "TP": 21,
}

# Column names as short, code-friendly labels for the report.
CATEGORY_LABELS = {
    "Cashless": "Cashless (1 day)",
    "Cashless Full case": "Cashless Full case (5 days)",
    "Benefit": "Benefit (5 days)",
    "FULL INVESTICATION": "FULL INVESTIGATION (7 days)",
    "MBV": "MBV (3 days)",
    "PRE CLAIM VERIFICATION": "PRE CLAIM VERIFICATION (3 days)",
    "Re Investigation": "Re Investigation (3 days)",
    "Reimbursement": "Reimbursement (5 days)",
    "Reimbursement-Half case": "Reimbursement-Half case (5 days)",
    "TP": "TP (21 days)",
}

MANAGER_COL = "Manager"
SUB_PRODUCT_COL = "Sub Product"
TAT_COL = "TAT"
MANAGER_CLOSED_DAYS_COL = "Manager closed days"
FINAL_CONCLUSION_COL = "Final Conclusion"
DISCREPANT_VALUE = "Discrepant"
MANAGER_CLOSED_DAYS_THRESHOLD = 2


def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Sheet1", header=0)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    df = df.dropna(subset=[MANAGER_COL, SUB_PRODUCT_COL])
    df = df[df[SUB_PRODUCT_COL].isin(STANDARD_TAT.keys())]
    return df


def build_report(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_standard_tat"] = df[SUB_PRODUCT_COL].map(STANDARD_TAT)
    df["_within_tat"] = df[TAT_COL] <= df["_standard_tat"]
    df["_discrepant"] = df[FINAL_CONCLUSION_COL] == DISCREPANT_VALUE
    df["_closed_2plus"] = df[MANAGER_CLOSED_DAYS_COL] >= MANAGER_CLOSED_DAYS_THRESHOLD

    managers = sorted(df[MANAGER_COL].dropna().unique())
    rows = []
    for manager in managers:
        sub = df[df[MANAGER_COL] == manager]
        row = {"Manager": manager}
        row["Total Cases"] = len(sub)
        row["Total Within TAT"] = int(sub["_within_tat"].sum())
        row["Total Discrepant"] = int(sub["_discrepant"].sum())
        row["Total Closed Days >= 2"] = int(sub["_closed_2plus"].sum())
        for category in STANDARD_TAT:
            cat_sub = sub[sub[SUB_PRODUCT_COL] == category]
            label = CATEGORY_LABELS[category]
            row[f"{label} - Cases"] = len(cat_sub)
            row[f"{label} - Within TAT"] = int(cat_sub["_within_tat"].sum())
            row[f"{label} - Discrepant"] = int(cat_sub["_discrepant"].sum())
            row[f"{label} - Closed Days >= 2"] = int(cat_sub["_closed_2plus"].sum())
        rows.append(row)

    report = pd.DataFrame(rows)

    # Grand total row
    total_row = {"Manager": "Grand Total"}
    total_row["Total Cases"] = report["Total Cases"].sum()
    total_row["Total Within TAT"] = report["Total Within TAT"].sum()
    total_row["Total Discrepant"] = report["Total Discrepant"].sum()
    total_row["Total Closed Days >= 2"] = report["Total Closed Days >= 2"].sum()
    for category in STANDARD_TAT:
        label = CATEGORY_LABELS[category]
        total_row[f"{label} - Cases"] = report[f"{label} - Cases"].sum()
        total_row[f"{label} - Within TAT"] = report[f"{label} - Within TAT"].sum()
        total_row[f"{label} - Discrepant"] = report[f"{label} - Discrepant"].sum()
        total_row[f"{label} - Closed Days >= 2"] = report[f"{label} - Closed Days >= 2"].sum()
    report = pd.concat([report, pd.DataFrame([total_row])], ignore_index=True)

    return report


def main():
    parser = argparse.ArgumentParser(description="Build Manager-wise TAT compliance report.")
    parser.add_argument("input", help="Path to the source .xlsx file")
    parser.add_argument("-o", "--output", help="Path to write the report .xlsx file")
    args = parser.parse_args()

    output = args.output
    if not output:
        if args.input.lower().endswith(".xlsx"):
            output = args.input[:-5] + "_Manager_TAT_Report.xlsx"
        else:
            output = args.input + "_Manager_TAT_Report.xlsx"

    df = load_data(args.input)
    report = build_report(df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report.to_excel(writer, sheet_name="Manager TAT Report", index=False)

    print(f"Report written to: {output}")
    print(f"Managers: {len(report) - 1}, Total cases: {report.iloc[-1]['Total Cases']}")


if __name__ == "__main__":
    sys.exit(main())
