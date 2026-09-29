"""
Streamlit app: upload an export, pick which report to build, download
the result(s) as Excel.

Covers three reports:
  - Health Open Cases: Client Wise + Manager Pending Closure
  - FO TAT: FO-wise TAT compliance from a "Cat Closed" export
  - Manager TAT: Manager-wise TAT compliance from a "Health Managers
    closed" export

Run locally with:
    streamlit run app.py
"""

from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

import fo_tat_report
import manager_tat_report
from health_open_cases_report import (
    build_client_report,
    build_manager_pending_closure_report,
)

st.set_page_config(page_title="SKD TAT Reports", page_icon="📋", layout="centered")


def to_excel_bytes(sheets: dict) -> bytes:
    """sheets: {sheet_name: DataFrame}"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()


st.title("SKD TAT Reports")

report_type = st.selectbox(
    "Which report do you want to build?",
    ["Health Open Cases", "FO TAT", "Manager TAT"],
)

if report_type == "Health Open Cases":
    st.write(
        "Upload a **Health open cases** export (.xlsx) to generate two reports: "
        "a Client-wise case breakdown, and a Manager Pending Closure report."
    )
    uploaded_file = st.file_uploader("Upload Health open cases file", type=["xlsx"], key="hoc")

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
            data=to_excel_bytes({"Client Wise": client_report}),
            file_name="Client_Wise_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("Manager Pending Closure")
        st.dataframe(manager_report, use_container_width=True)
        st.download_button(
            "Download Manager Pending Closure report",
            data=to_excel_bytes({"Manager Pending Closure": manager_report}),
            file_name="Manager_Pending_Closure_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Waiting for a file to be uploaded.")

elif report_type == "FO TAT":
    st.write(
        "Upload a **Cat Closed** export (.xlsx) to generate the FO-wise TAT "
        "compliance report (cases, within-TAT, and Discrepant counts per FO "
        "and category)."
    )
    uploaded_file = st.file_uploader("Upload Cat Closed file", type=["xlsx"], key="fo_tat")

    if uploaded_file is not None:
        try:
            df = fo_tat_report.load_data(uploaded_file)
        except Exception as exc:
            st.error(f"Could not read the file: {exc}")
            st.stop()

        report = fo_tat_report.build_report(df)

        st.success(f"Processed {len(df)} rows across {len(report) - 1} FOs.")
        st.dataframe(report, use_container_width=True)
        st.download_button(
            "Download FO TAT report",
            data=to_excel_bytes({"FO TAT Report": report}),
            file_name="FO_TAT_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Waiting for a file to be uploaded.")

elif report_type == "Manager TAT":
    st.write(
        "Upload a **Health Managers closed** export (.xlsx) to generate the "
        "Manager-wise TAT compliance report."
    )
    uploaded_file = st.file_uploader("Upload Health Managers closed file", type=["xlsx"], key="mgr_tat")

    if uploaded_file is not None:
        try:
            df = manager_tat_report.load_data(uploaded_file)
        except Exception as exc:
            st.error(f"Could not read the file: {exc}")
            st.stop()

        report = manager_tat_report.build_report(df)

        st.success(f"Processed {len(df)} rows across {len(report) - 1} managers.")
        st.dataframe(report, use_container_width=True)
        st.download_button(
            "Download Manager TAT report",
            data=to_excel_bytes({"Manager TAT Report": report}),
            file_name="Manager_TAT_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Waiting for a file to be uploaded.")
