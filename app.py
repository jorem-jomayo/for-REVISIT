import streamlit as st
import pandas as pd
import io
from datetime import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# -------------------------------------------------------------------
# Helper: clean text (same as VBA CleanText)
# -------------------------------------------------------------------
def clean_text(val):
    if pd.isna(val):
        return ""
    if isinstance(val, str):
        # remove non-printable characters and trim
        import re
        s = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', val)  # remove control chars
        s = s.strip()
        # replace multiple spaces
        while "  " in s:
            s = s.replace("  ", " ")
        return s
    return str(val).strip()

# -------------------------------------------------------------------
# Main classification function
# -------------------------------------------------------------------
def classify_rows(df):
    """
    Takes a DataFrame with columns: "DL TYPE", "ACCOUNT TYPE", "REMARKS"
    Returns a dict with sheet names as keys and DataFrames as values.
    """
    # Clean the three key columns
    df["DL TYPE"] = df["DL TYPE"].apply(clean_text).str.upper()
    df["ACCOUNT TYPE"] = df["ACCOUNT TYPE"].apply(clean_text).str.upper()
    df["REMARKS"] = df["REMARKS"].apply(clean_text).str.upper()

    # Initialize result dictionary
    sheets = {
        "DL1": [],
        "DL2": [],
        "DL3": [],
        "DL4 CTS": [],
        "DL4 CTS REV": [],
        "DL4 REM": [],
        "DL5": [],
        "DL6": []
    }

    total = len(df)
    progress_bar = st.progress(0, text="Processing rows...")
    status_text = st.empty()

    for idx, (_, row) in enumerate(df.iterrows()):
        dl_type = row["DL TYPE"]
        acc_type = row["ACCOUNT TYPE"]
        remarks = row["REMARKS"]

        dest = None
        if dl_type == "DL1":
            dest = "DL1"
        elif dl_type == "DL2":
            dest = "DL2"
        elif dl_type == "DL3":
            dest = "DL3"
        elif dl_type == "DL4":
            if acc_type == "CYCLE 3 CTS":
                if remarks == "NEW ENDO":
                    dest = "DL4 CTS"
                elif remarks == "REVISIT":
                    dest = "DL4 CTS REV"
            elif acc_type == "CYCLE 3 REM":
                dest = "DL4 REM"
        elif dl_type == "DL5":
            dest = "DL5"
        elif dl_type == "DL6":
            dest = "DL6"

        if dest is not None:
            sheets[dest].append(row)

        # Update progress
        if (idx + 1) % 10 == 0 or idx + 1 == total:
            progress = (idx + 1) / total
            progress_bar.progress(progress, text=f"Processing row {idx+1} of {total}")
            status_text.text(f"Processed {idx+1} rows...")

    progress_bar.empty()
    status_text.empty()

    # Convert lists to DataFrames
    for key in sheets:
        if sheets[key]:
            sheets[key] = pd.DataFrame(sheets[key])
        else:
            sheets[key] = pd.DataFrame()  # empty DataFrame

    return sheets

# -------------------------------------------------------------------
# Streamlit UI
# -------------------------------------------------------------------
st.set_page_config(page_title="PEN RATE Splitter", layout="wide")
st.title("📊 PEN RATE – Split by DL TYPE")

uploaded_file = st.file_uploader("Upload your PEN RATE.xlsx file", type=["xlsx"])

if uploaded_file is not None:
    # Read the Excel file
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")
        st.success("File loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    # Check required columns
    required_cols = ["DL TYPE", "ACCOUNT TYPE", "REMARKS", "ENDO DATE"]
    missing = [col for col in required_cols if col not in df_raw.columns]
    if missing:
        st.error(f"Missing columns: {missing}")
        st.stop()

    # Convert ENDO DATE to datetime (if not already)
    df_raw["ENDO DATE"] = pd.to_datetime(df_raw["ENDO DATE"], errors="coerce")

    # Create a month-year column for filtering
    df_raw["MONTH_YEAR"] = df_raw["ENDO DATE"].dt.strftime("%B %Y")  # e.g., "July 2026"

    # Drop rows with invalid dates
    df_valid = df_raw.dropna(subset=["ENDO DATE"]).copy()

    if df_valid.empty:
        st.warning("No valid ENDO DATE found. Please check your data.")
        st.stop()

    # Get unique month-year values, sorted chronologically
    unique_months = sorted(df_valid["MONTH_YEAR"].unique(), key=lambda x: datetime.strptime(x, "%B %Y"))
    selected_month = st.selectbox("Select ENDO DATE month", unique_months)

    # Filter data by selected month
    df_filtered = df_valid[df_valid["MONTH_YEAR"] == selected_month].copy()

    if df_filtered.empty:
        st.info(f"No data for {selected_month}.")
        st.stop()

    st.write(f"**Rows for {selected_month}:** {len(df_filtered)}")

    # Button to trigger splitting
    if st.button("Split Data"):
        with st.spinner("Splitting data..."):
            sheets = classify_rows(df_filtered)

        st.success("Splitting complete!")

        # Display each sheet in an expander
        for sheet_name, df_sheet in sheets.items():
            with st.expander(f"📄 {sheet_name}  ({len(df_sheet)} rows)"):
                if df_sheet.empty:
                    st.write("No data.")
                else:
                    st.dataframe(df_sheet, use_container_width=True)

        # Offer download as single Excel file with multiple sheets
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet_name, df_sheet in sheets.items():
                if not df_sheet.empty:
                    df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    # Write an empty sheet with just a note
                    pd.DataFrame({"Note": ["No data"]}).to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)

        st.download_button(
            label="📥 Download Split File (Excel)",
            data=output,
            file_name=f"PEN_RATE_Split_{selected_month.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload an Excel file to begin.")