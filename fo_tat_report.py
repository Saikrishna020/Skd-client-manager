"""
FO-wise TAT compliance report.

Reads the "Sheet1" data from the SKD Corinsoft export (header on row 2,
data from row 3), groups cases by FO Name and Sub Product, and reports,
for each FO:
    - Total cases handled
    - Total cases completed within TAT
    - Per category: case count, within-TAT count, and Discrepant count
      (Cashless, Cashless Full case, Benefit, FULL INVESTICATION, MBV,
       PRE CLAIM VERIFICATION, Re Investigation, Reimbursement,
       Reimbursement-Half case, TP)

"Within TAT" = FO Completed TAT (days) <= the standard TAT for that
sub product (defined in STANDARD_TAT below).

"Discrepant" = Final Conclusion column equals "Discrepant".

Usage:
    python fo_tat_report.py "Cat Closed 01 Apr to 31 Aug 26.xlsx" -o "FO_TAT_Report.xlsx"

If -o is omitted, the output file is named "<input>_FO_TAT_Report.xlsx".
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

FO_NAME_COL = "FO Name"
SUB_PRODUCT_COL = "Sub Product"
FO_TAT_COL = "FO Completed TAT"
CREATED_DATE_COL = "Created Date"
FO_COMPLETED_DATE_COL = "FO Completed Date"
FINAL_CONCLUSION_COL = "Final Conclusion"
DISCREPANT_VALUE = "Discrepant"


def _read_with_header_detection(path: str) -> pd.DataFrame:
    """Some exports have a title row above the header, some don't. Try
    header=0 first, and fall back to header=1 (skipping a title row) if
    the expected columns aren't found."""
    df = pd.read_excel(path, sheet_name="Sheet1", header=0)
    if FO_NAME_COL in df.columns and SUB_PRODUCT_COL in df.columns:
        return df
    df = pd.read_excel(path, sheet_name="Sheet1", header=1)
    if FO_NAME_COL in df.columns and SUB_PRODUCT_COL in df.columns:
        return df
    raise ValueError(
        f"Could not find '{FO_NAME_COL}' / '{SUB_PRODUCT_COL}' columns "
        "with header on row 1 or row 2. Check the source file's layout."
    )


def load_data(path: str) -> pd.DataFrame:
    df = _read_with_header_detection(path)
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

    if FO_TAT_COL not in df.columns:
        # Some exports don't have a precomputed TAT column - derive it
        # from the completion and creation dates instead.
        created = pd.to_datetime(df[CREATED_DATE_COL], dayfirst=True, errors="coerce")
        completed = pd.to_datetime(df[FO_COMPLETED_DATE_COL], dayfirst=True, errors="coerce")
        df[FO_TAT_COL] = (completed - created).dt.days

    df = df.dropna(subset=[FO_NAME_COL, SUB_PRODUCT_COL])
    df = df[df[SUB_PRODUCT_COL].isin(STANDARD_TAT.keys())]
    return df


def build_report(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_standard_tat"] = df[SUB_PRODUCT_COL].map(STANDARD_TAT)
    df["_within_tat"] = df[FO_TAT_COL] <= df["_standard_tat"]
    df["_discrepant"] = df[FINAL_CONCLUSION_COL] == DISCREPANT_VALUE

    fo_names = sorted(df[FO_NAME_COL].dropna().unique())
    rows = []
    for fo in fo_names:
        sub = df[df[FO_NAME_COL] == fo]
        row = {"FO Name": fo}
        row["Total Cases"] = len(sub)
        row["Total Within TAT"] = int(sub["_within_tat"].sum())
        row["Total Discrepant"] = int(sub["_discrepant"].sum())
        for category in STANDARD_TAT:
            cat_sub = sub[sub[SUB_PRODUCT_COL] == category]
            label = CATEGORY_LABELS[category]
            row[f"{label} - Cases"] = len(cat_sub)
            row[f"{label} - Within TAT"] = int(cat_sub["_within_tat"].sum())
            row[f"{label} - Discrepant"] = int(cat_sub["_discrepant"].sum())
        rows.append(row)

    report = pd.DataFrame(rows)

    # Grand total row
    total_row = {"FO Name": "Grand Total"}
    total_row["Total Cases"] = report["Total Cases"].sum()
    total_row["Total Within TAT"] = report["Total Within TAT"].sum()
    total_row["Total Discrepant"] = report["Total Discrepant"].sum()
    for category in STANDARD_TAT:
        label = CATEGORY_LABELS[category]
        total_row[f"{label} - Cases"] = report[f"{label} - Cases"].sum()
        total_row[f"{label} - Within TAT"] = report[f"{label} - Within TAT"].sum()
        total_row[f"{label} - Discrepant"] = report[f"{label} - Discrepant"].sum()
    report = pd.concat([report, pd.DataFrame([total_row])], ignore_index=True)

    return report


def main():
    parser = argparse.ArgumentParser(description="Build FO-wise TAT compliance report.")
    parser.add_argument("input", help="Path to the source .xlsx file")
    parser.add_argument("-o", "--output", help="Path to write the report .xlsx file")
    args = parser.parse_args()

    output = args.output
    if not output:
        if args.input.lower().endswith(".xlsx"):
            output = args.input[:-5] + "_FO_TAT_Report.xlsx"
        else:
            output = args.input + "_FO_TAT_Report.xlsx"

    df = load_data(args.input)
    report = build_report(df)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report.to_excel(writer, sheet_name="FO TAT Report", index=False)

    print(f"Report written to: {output}")
    print(f"FOs: {len(report) - 1}, Total cases: {report.iloc[-1]['Total Cases']}")


if __name__ == "__main__":
    sys.exit(main())
