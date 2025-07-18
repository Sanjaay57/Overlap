import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="MS Number Overlap Checker", layout="wide")
st.title("🔍 MS Number Overlap Checker")
st.markdown("Compare a student list against a main database and get full details if matched.")

uploaded_file = st.file_uploader("📤 Upload Excel File (with at least 2 sheets)", type=["xlsx"])

if uploaded_file:
    try:
        # Load all sheets
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        sheet_names = list(all_sheets.keys())

        # User selection for main & compare sheet
        col1, col2 = st.columns(2)
        with col1:
            main_sheet_name = st.selectbox("📚 Select Main Sheet (with full student details)", sheet_names)
        with col2:
            compare_sheet_name = st.selectbox("📝 Select Compare Sheet (with MS numbers)", sheet_names)

        # Proceed only if different sheets selected
        if main_sheet_name == compare_sheet_name:
            st.warning("⚠️ Please select two different sheets.")
            st.stop()

        if st.button("🔎 Compare"):
            # Load sheets
            main_df = all_sheets[main_sheet_name].copy()
            compare_df = all_sheets[compare_sheet_name].copy()

            # Clean column headers
            main_df.columns = main_df.columns.str.strip()
            compare_df.columns = compare_df.columns.str.strip()

            # Identify MS number column in each sheet
            ms_col_main = main_df.columns[1]  # assume 2nd column in main
            ms_col_compare = compare_df.columns[0]  # assume 1st column in compare

            # Normalize MS numbers
            main_df[ms_col_main] = main_df[ms_col_main].apply(lambda x: str(int(x)) if isinstance(x, float) else str(x).strip())
            compare_df[ms_col_compare] = compare_df[ms_col_compare].apply(lambda x: str(int(x)) if isinstance(x, float) else str(x).strip())

            # Lookup setup
            lookup_df = main_df.set_index(ms_col_main)

            # Fields to extract
            expected_fields = ['MS Number', 'Student Name', 'District', 'Institution Name', 'Campus', 'Course']
            column_map = {
                'MS Number': ms_col_main,
                'Student Name': 'Student Name',
                'District': 'District',
                'Institution Name': 'Institution Name',
                'Campus': 'Campus',
                'Course': 'Course'
            }

            # Check all required fields exist
            missing = [v for k, v in column_map.items() if v not in main_df.columns]
            if missing:
                st.error(f"❌ Missing columns in main sheet: {missing}")
                st.stop()

            # Build result
            result = []
            for _, row in compare_df.iterrows():
                ms_number = row[ms_col_compare]
                if ms_number in lookup_df.index:
                    match = lookup_df.loc[ms_number]
                    result.append({
                        "MS Number": ms_number,
                        "Student Name": match.get(column_map['Student Name'], ""),
                        "District": match.get(column_map['District'], ""),
                        "Institution Name": match.get(column_map['Institution Name'], ""),
                        "Campus": match.get(column_map['Campus'], ""),
                        "Course": match.get(column_map['Course'], ""),
                        "Status": "Overlapped"
                    })
                else:
                    result.append({
                        "MS Number": ms_number,
                        "Student Name": "",
                        "District": "",
                        "Institution Name": "",
                        "Campus": "",
                        "Course": "",
                        "Status": "Unique"
                    })

            result_df = pd.DataFrame(result)
            result_df.index = range(1, len(result_df) + 1)

            st.success("✅ Comparison complete.")
            st.dataframe(result_df, use_container_width=True)

            # Export to Excel
            output = BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "MS_Compare_Result"

            for r in dataframe_to_rows(result_df.reset_index(), index=False, header=True):
                ws.append(r)

            for col in ws.columns:
                max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_len + 3

            wb.save(output)

            st.download_button(
                label="📥 Download Result Excel",
                data=output.getvalue(),
                file_name="MS_Number_Comparison_Result.xlsx"
            )

    except Exception as e:
        st.error(f"❌ Something went wrong: {e}")
