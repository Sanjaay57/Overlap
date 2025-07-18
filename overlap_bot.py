import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# === Page Setup ===
st.set_page_config(page_title="MS Number Overlap Checker", layout="wide")
st.markdown("<h1 style='text-align: center;'>MS Number Overlap Check</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>TN Model Schools</h4>", unsafe_allow_html=True)

# === File Upload ===
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    uploaded_file = st.file_uploader("📤 Upload Excel file with Main & Sheet1", type=["xlsx"])

if not uploaded_file:
    st.info("📁 Please upload the Excel file to continue.")
    st.stop()

try:
    # Read all sheets
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    if "Main" not in all_sheets or "Sheet1" not in all_sheets:
        st.error("❌ Excel must contain two sheets: 'Main' and 'Sheet1'")
        st.stop()

    main_df = all_sheets["Main"].copy()
    sheet1_df = all_sheets["Sheet1"].copy()

    # Clean column names
    main_df.columns = main_df.columns.str.strip()
    sheet1_df.columns = sheet1_df.columns.str.strip()

    # MS number assumed in second column of Main and first column of Sheet1
    ms_col_main = main_df.columns[1]
    ms_col_sheet1 = sheet1_df.columns[0]

    # Clean MS numbers to string format
    main_df[ms_col_main] = main_df[ms_col_main].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )
    sheet1_df[ms_col_sheet1] = sheet1_df[ms_col_sheet1].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )

    # Lookup data from Main sheet
    main_lookup = main_df.set_index(ms_col_main)

    result_rows = []
    for _, row in sheet1_df.iterrows():
        ms_number = row[ms_col_sheet1]
        if ms_number in main_lookup.index:
            data = main_lookup.loc[ms_number]
            result_rows.append({
                "MS Number": ms_number,
                "District": data.get("District", ""),
                "Student Name": data.get("Student Name", ""),
                "Institution Name": data.get("Institution Name", ""),
                "Campus": data.get("Campus", ""),
                "Course": data.get("Course", ""),
                "Status": "Overlapped"
            })
        else:
            result_rows.append({
                "MS Number": ms_number,
                "District": row.get("District", ""),
                "Student Name": row.get("Student Name", ""),
                "Institution Name": row.get("Institution Name", ""),
                "Campus": row.get("Campus", ""),
                "Course": row.get("Course", ""),
                "Status": "Unique"
            })

    # Create result DataFrame
    result_df = pd.DataFrame(result_rows)
    result_df.index = range(1, len(result_df) + 1)

    st.success(f"✅ Checked {len(result_df)} MS Numbers from 'Sheet1' against 'Main'")
    st.dataframe(result_df, use_container_width=True)

    # === Excel Export ===
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MS Overlap Check"

    for r in dataframe_to_rows(result_df.reset_index(), index=False, header=True):
        ws.append(r)

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    wb.save(output)

    st.download_button(
        label="📥 Download Result as Excel",
        data=output.getvalue(),
        file_name="MS_Overlap_Result.xlsx"
    )

except Exception as e:
    st.error(f"❌ Error: {e}")
