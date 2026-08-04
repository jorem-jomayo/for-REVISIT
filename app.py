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
    progress_bar = st.progress(0, text="Classifying rows...")
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
            progress_bar.progress(progress, text=f"Classifying row {idx+1} of {total}")
            status_text.text(f"Classified {idx+1} rows...")

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
# New function: copy NOT MATCH rows to NEW DL TYPE, add destination column
# Now with optional inclusion of DL6 matching against DL1
# -------------------------------------------------------------------
def add_match_and_copy(sheets, include_dl6=False):
    """
    Modifies sheets in place:
      - Adds column "MATCH_STATUS" to source sheets (DL1, DL2, DL4 CTS, DL4 REM, DL5)
      - Adds column "FOR_NEW_DL_TYPE" with destination label for NOT MATCH rows
      - Copies ONLY the NOT MATCH rows into a new sheet "NEW DL TYPE"
      - Original rows stay in their source sheets (no deletion)
      - If include_dl6 is True, also processes DL6 (matches against DL1)
    """
    # Base rules: (source, target, destination_label_for_not_match)
    rules = [
        ("DL1", "DL2", "FOR DL2"),
        ("DL2", "DL3", "FOR DL3"),
        ("DL4 CTS", "DL4 CTS REV", "FOR DL4"),
        ("DL4 REM", "DL5", "FOR DL5"),
        ("DL5", "DL6", "FOR DL6")
    ]

    # Optionally add DL6 rule (match DL6 against DL1)
    if include_dl6:
        rules.append(("DL6", "DL1", "FOR DL1"))

    # First, compute total rows that will be processed across all source sheets
    total_rows = 0
    for src, _, _ in rules:
        if src in sheets:
            total_rows += len(sheets[src])
    if total_rows == 0:
        # Nothing to match
        sheets["NEW DL TYPE"] = pd.DataFrame()
        return

    not_matched_rows = []  # will hold DataFrames of NOT MATCH rows (to copy)

    progress_bar = st.progress(0, text="Matching rows...")
    status_text = st.empty()
    processed = 0

    for src, tgt, dest_label in rules:
        if src not in sheets or tgt not in sheets:
            continue
        df_src = sheets[src]
        df_tgt = sheets[tgt]
        if df_src.empty:
            processed += len(df_src)  # 0
            continue

        # If target is empty, all rows are NOT MATCH
        if df_tgt.empty:
            df_src["MATCH_STATUS"] = "NOT MATCH"
            df_src["FOR_NEW_DL_TYPE"] = dest_label   # all rows get this label
            # Copy all rows as NOT MATCH
            not_matched_rows.append(df_src.copy())
            sheets[src] = df_src
            processed += len(df_src)
            progress = processed / total_rows
            progress_bar.progress(progress, text=f"Matching rows... {processed}/{total_rows}")
            status_text.text(f"Matched {processed} rows...")
            continue

        # Get values from the first column (column A)
        src_vals = df_src.iloc[:, 0]
        tgt_vals = set(df_tgt.iloc[:, 0].dropna().unique())

        # Add match status
        df_src["MATCH_STATUS"] = src_vals.apply(lambda x: "MATCH" if x in tgt_vals else "NOT MATCH")
        # Add destination column: for NOT MATCH rows, set to dest_label; for MATCH rows, set to ""
        df_src["FOR_NEW_DL_TYPE"] = df_src["MATCH_STATUS"].apply(
            lambda status: dest_label if status == "NOT MATCH" else ""
        )

        # Copy only NOT MATCH rows
        not_matched = df_src[df_src["MATCH_STATUS"] == "NOT MATCH"]
        if not not_matched.empty:
            not_matched_rows.append(not_matched.copy())

        # Keep the source sheet as is (with both new columns)
        sheets[src] = df_src

        # Update progress
        processed += len(df_src)
        progress = min(processed / total_rows, 1.0)
        progress_bar.progress(progress, text=f"Matching rows... {processed}/{total_rows}")
        status_text.text(f"Matched {processed} rows...")

    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()

    # Create NEW DL TYPE sheet from all copied NOT MATCH rows
    if not_matched_rows:
        new_df = pd.concat(not_matched_rows, ignore_index=True)
        sheets["NEW DL TYPE"] = new_df
    else:
        sheets["NEW DL TYPE"] = pd.DataFrame()

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

    # ---- New: checkbox to include DL6 in matching ----
    include_dl6 = st.checkbox("Include DL6 in matching (match DL6 against DL1)")

    # Button to trigger splitting
    if st.button("Split Data"):
        with st.spinner("Processing..."):
            sheets = classify_rows(df_filtered)
            # Apply matching and copy NOT MATCH rows, passing the checkbox value
            add_match_and_copy(sheets, include_dl6=include_dl6)

        st.success("Processing complete!")

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
