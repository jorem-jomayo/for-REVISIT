import streamlit as st
import pandas as pd
import io
from datetime import datetime
import re

# -------------------------------------------------------------------
# Helper: clean text (removes hidden chars and extra spaces)
# -------------------------------------------------------------------
def clean_text(val):
    if pd.isna(val):
        return ""
    if isinstance(val, str):
        s = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', val)
        s = s.strip()
        while "  " in s:
            s = s.replace("  ", " ")
        return s
    return str(val).strip()

# -------------------------------------------------------------------
# Add MATCHING column (Excel-style MATCH)
# -------------------------------------------------------------------
def add_matching_columns(sheets):
    match_rules = [
        ("DL1", "DL2"),
        ("DL2", "DL3"),
        ("DL4 CTS", "DL4 CTS REV"),
        ("DL4 REM", "DL5"),
        ("DL5", "DL6"),
    ]

    for src_name, tgt_name in match_rules:
        src_df = sheets.get(src_name)
        tgt_df = sheets.get(tgt_name)

        if src_df is None or tgt_df is None:
            continue
        if src_df.empty or tgt_df.empty:
            if not src_df.empty:
                src_df["MATCHING"] = "Target sheet empty"
            continue

        # Use the first column as the key (like column A in Excel)
        key_col = src_df.columns[0]
        tgt_key_col = tgt_df.columns[0]

        # Map target values to Excel row numbers (index + 2)
        match_map = dict(zip(tgt_df[tgt_key_col], tgt_df.index + 2))

        src_df["MATCHING"] = src_df[key_col].map(match_map)
        src_df["MATCHING"] = src_df["MATCHING"].fillna("Not found").astype(str)

    # For sheets without a matching rule, add an empty MATCHING column
    for name, df in sheets.items():
        if "MATCHING" not in df.columns and not df.empty:
            df["MATCHING"] = "No rule"

    return sheets

# -------------------------------------------------------------------
# Main classification function (splits into sheets)
# -------------------------------------------------------------------
def classify_rows(df):
    df_clean = df.copy()
    df_clean["DL TYPE"] = df_clean["DL TYPE"].apply(clean_text).str.upper()
    df_clean["ACCOUNT TYPE"] = df_clean["ACCOUNT TYPE"].apply(clean_text).str.upper()
    df_clean["REMARKS"] = df_clean["REMARKS"].apply(clean_text).str.upper()

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

    total = len(df_clean)
    progress_bar = st.progress(0, text="Starting...")
    status_text = st.empty()

    for idx, (_, row) in enumerate(df_clean.iterrows()):
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
            sheets[key] = pd.DataFrame()

    # Add MATCHING columns
    sheets = add_matching_columns(sheets)

    return sheets

# -------------------------------------------------------------------
# Streamlit UI
# -------------------------------------------------------------------
st.set_page_config(page_title="PEN RATE Splitter", layout="wide")
st.title("📊 PEN RATE – Split by DL TYPE + MATCHING column")

uploaded_file = st.file_uploader("Upload your PEN RATE.xlsx file", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")
        st.success("File loaded successfully!")
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    required_cols = ["DL TYPE", "ACCOUNT TYPE", "REMARKS", "ENDO DATE"]
    missing = [col for col in required_cols if col not in df_raw.columns]
    if missing:
        st.error(f"Missing columns: {missing}")
        st.stop()

    df_raw["ENDO DATE"] = pd.to_datetime(df_raw["ENDO DATE"], errors="coerce")
    df_raw["MONTH_YEAR"] = df_raw["ENDO DATE"].dt.strftime("%B %Y")

    df_valid = df_raw.dropna(subset=["ENDO DATE"]).copy()

    if df_valid.empty:
        st.warning("No valid ENDO DATE found.")
        st.stop()

    unique_months = sorted(df_valid["MONTH_YEAR"].unique(), key=lambda x: datetime.strptime(x, "%B %Y"))
    selected_month = st.selectbox("Select ENDO DATE month", unique_months)

    df_filtered = df_valid[df_valid["MONTH_YEAR"] == selected_month].copy()

    if df_filtered.empty:
        st.info(f"No data for {selected_month}.")
        st.stop()

    st.write(f"**Rows for {selected_month}:** {len(df_filtered)}")

    if st.button("Split Data"):
        with st.spinner("Splitting data..."):
            sheets = classify_rows(df_filtered)

        st.success("Splitting complete!")

        # ---------- Show each sheet in its own expander ----------
        for sheet_name, df_sheet in sheets.items():
            with st.expander(f"📄 {sheet_name}  ({len(df_sheet)} rows)"):
                if df_sheet.empty:
                    st.write("No data.")
                else:
                    st.dataframe(df_sheet, use_container_width=True)

        # ---------- Combine all sheets into one CSV ----------
        all_data = []
        for sheet_name, df_sheet in sheets.items():
            if not df_sheet.empty:
                df_copy = df_sheet.copy()
                df_copy.insert(0, "SHEET", sheet_name)   # Add sheet name as first column
                all_data.append(df_copy)

        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)

            # Convert to CSV
            csv_buffer = io.StringIO()
            combined_df.to_csv(csv_buffer, index=False)
            csv_data = csv_buffer.getvalue()

            st.download_button(
                label="📥 Download all data as one CSV",
                data=csv_data,
                file_name=f"PEN_RATE_Split_{selected_month.replace(' ', '_')}.csv",
                mime="text/csv"
            )

            st.info("ℹ️ CSV is plain text – font size is controlled by your viewing software (e.g., Excel, Notepad).")
        else:
            st.warning("No data to export.")

else:
    st.info("Please upload an Excel file to begin.")