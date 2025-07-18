import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# === UI Setup ===
st.set_page_config(page_title="TN Model Schools MS Number Check", layout="wide")
st.markdown("<h1 style='text-align: center;'>TN Model Schools MS Number Verification</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>MS CG Team</h4>", unsafe_allow_html=True)

# === Centered File Upload ===
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    uploaded_file = st.file_uploader("📤 Upload Excel File (with main sheet & new MS numbers)", type=["xlsx"])

# === Centered Info Box if no file uploaded ===
if not uploaded_file:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("📁 Please upload the Excel file to continue.")
    st.stop()

try:
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    sheet_names = list(all_sheets.keys())

    # === Sidebar Selection ===
    st.sidebar.header("🔍 Sheet Selection")
    main_sheet = st.sidebar.selectbox("📘 Main Sheet (All Joined Students)", sheet_names)
    new_sheet = st.sidebar.selectbox("🆕 Sheet with New MS Numbers", [s for s in sheet_names if s != main_sheet])

    if st.sidebar.button("🔎 Check for Missing MS Numbers"):
        main_df = all_sheets[main_sheet].copy()
        new_df = all_sheets[new_sheet].copy()

        # Clean columns and MS Numbers
        main_df.columns = main_df.columns.str.strip()
        new_df.columns = new_df.columns.str.strip()

        main_ms_col = main_df.columns[1]  # Assuming MS Number is 2nd column
        new_ms_col = new_df.columns[0]    # Assuming MS Number is 1st column

        main_df[main_ms_col] = main_df[main_ms_col].apply(
            lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
        )
        new_df[new_ms_col] = new_df[new_ms_col].apply(
            lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
        )

        # Compare MS numbers
        main_ms_set = set(main_df[main_ms_col].dropna())
        new_ms_set = set(new_df[new_ms_col].dropna())
        missing_ms = new_ms_set - main_ms_set

        # Prepare Missing MS Table
        missing_df = new_df[new_df[new_ms_col].isin(missing_ms)].copy()
        missing_df.rename(columns={new_ms_col: "MS Number"}, inplace=True)

        # Prepare lookup from main sheet for student info
        lookup_df = main_df.set_index(main_ms_col)
        get_value = lambda x, col: lookup_df[col][x] if x in lookup_df.index and col in lookup_df else "❌ Not Found"

        # Add Info Columns
        missing_df["Student Name"] = missing_df["MS Number"].apply(lambda x: get_value(x, "Student Name"))
        missing_df["District"] = missing_df["MS Number"].apply(lambda x: get_value(x, "District"))
        missing_df["Institution Name"] = missing_df["MS Number"].apply(lambda x: get_value(x, "Institution Name"))
        missing_df["Campus"] = missing_df["MS Number"].apply(lambda x: get_value(x, "Campus"))
        missing_df["Course"] = missing_df["MS Number"].apply(lambda x: get_value(x, "Course"))

        # Final Display
        display_df = missing_df[["MS Number", "Student Name", "District", "Institution Name", "Campus", "Course"]]
        display_df.index = range(1, len(display_df) + 1)

        st.warning(f"⚠️ Found {len(display_df)} MS Number(s) not present in the main sheet.")
        st.dataframe(display_df, use_container_width=True)

        # === Excel Download ===
        output = BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Missing MS Numbers"

        for r in dataframe_to_rows(display_df.reset_index(), index=False, header=True):
            ws.append(r)

        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2

        wb.save(output)

        st.download_button(
            "📥 Download Missing MS Numbers",
            data=output.getvalue(),
            file_name="missing_ms_numbers.xlsx"
        )

except Exception as e:
    st.error(f"❌ Error: {e}")
