import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# === UI Setup ===
st.set_page_config(page_title="World Tab MS Overlap Checker", layout="wide")
st.markdown("<h1 style='text-align: center;'>World Tab MS Number Check</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>TN Model Schools</h4>", unsafe_allow_html=True)

# === File Upload ===
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    uploaded_file = st.file_uploader("📤 Upload Excel File (must contain 'Main' & 'World')", type=["xlsx"])

if not uploaded_file:
    st.info("📁 Please upload the Excel file to continue.")
    st.stop()

try:
    # Load all sheets
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    if "Main" not in all_sheets or "World" not in all_sheets:
        st.error("❌ Excel must contain two sheets: 'Main' and 'World'")
        st.stop()

    main_df = all_sheets["Main"].copy()
    world_df = all_sheets["World"].copy()

    # Clean column names
    main_df.columns = main_df.columns.str.strip()
    world_df.columns = world_df.columns.str.strip()

    # Assume MS Number in Main is column 1; in World it’s column 0
    ms_col_main = main_df.columns[1]
    ms_col_world = world_df.columns[0]

    # Convert MS Numbers to string
    main_df[ms_col_main] = main_df[ms_col_main].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )
    world_df[ms_col_world] = world_df[ms_col_world].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
    )

    # Lookup table from Main sheet
    lookup_df = main_df.set_index(ms_col_main)

    # Prepare result rows
    result = []
    for _, row in world_df.iterrows():
        ms_number = row[ms_col_world]
        if ms_number in lookup_df.index:
            matched = lookup_df.loc[ms_number]
            result.append({
                "MS Number": ms_number,
                "Student Name": matched.get("Student Name", ""),
                "District": matched.get("District", ""),
                "Institution Name": matched.get("Institution Name", ""),
                "Campus": matched.get("Campus", ""),
                "Course": matched.get("Course", ""),
                "World App Status": "Overlapped"
            })
        else:
            result.append({
                "MS Number": ms_number,
                "Student Name": "",
                "District": "",
                "Institution Name": "",
                "Campus": "",
                "Course": "",
                "World App Status": "Unique"
            })

    # Convert to DataFrame
    result_df = pd.DataFrame(result)
    result_df.index = range(1, len(result_df) + 1)

    st.success("✅ World tab successfully compared with Main sheet.")
    st.dataframe(result_df, use_container_width=True)

    # === Excel Download ===
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "World_Overlap_Check"

    for r in dataframe_to_rows(result_df.reset_index(), index=False, header=True):
        ws.append(r)

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    wb.save(output)

    st.download_button(
        label="📥 Download Result as Excel",
        data=output.getvalue(),
        file_name="World_MS_Overlap_Result.xlsx"
    )

except Exception as e:
    st.error(f"❌ Error processing file: {e}")
