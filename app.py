"""
Streamlit app: upload a "Health open cases" export and download two
reports - a Client-wise breakdown and a Manager Pending Closure report.

Run locally with:
    streamlit run app.py
"""

from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

from health_open_cases_report import (
    build_client_report,
    build_manager_pending_closure_report,
)

st.set_page_config(page_title="Health Open Cases Report", page_icon="📋", layout="centered")

st.title("Health Open Cases Report")
st.write(
    "Upload a **Health open cases** export (.xlsx) to generate two reports: "
    "a Client-wise case breakdown, and a Manager Pending Closure report."
)

uploaded_file = st.file_uploader("Upload Health open cases file", type=["xlsx"])


def to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()


if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1", header=0)
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")
        st.stop()

    required_cols = {"Client", "Sub Product", "Manager", "Status", "SKD TAT-D", "CAT Completed Date"}
    missing = required_cols - set(df.columns)
    if missing:
        st.error(f"The uploaded file is missing expected column(s): {', '.join(sorted(missing))}")
        st.stop()

    today = datetime.now()

    client_report = build_client_report(df)
    manager_report = build_manager_pending_closure_report(df, today=today)

    st.success(f"Processed {len(df)} rows. Report date: {today.strftime('%d %b %Y')}")

    st.subheader("Client Wise")
    st.dataframe(client_report, use_container_width=True)
    st.download_button(
        "Download Client Wise report",
        data=to_excel_bytes(client_report, "Client Wise"),
        file_name="Client_Wise_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Manager Pending Closure")
    st.dataframe(manager_report, use_container_width=True)
    st.download_button(
        "Download Manager Pending Closure report",
        data=to_excel_bytes(manager_report, "Manager Pending Closure"),
        file_name="Manager_Pending_Closure_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Waiting for a file to be uploaded.")
