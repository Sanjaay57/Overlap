import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# === Page Setup ===
st.set_page_config(page_title="TN Model Schools Overlap Bot", layout="wide")
st.markdown("<h1 style='text-align: center;'>TN Model Schools Student Overlap</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>MS CG Team</h4>", unsafe_allow_html=True)

# === Centered Search Box ===
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    search_query = st.text_input("🔍 Search by EMIS / Name", placeholder="Enter EMIS number or Name")

# === Centered File Upload ===
col4, col5, col6 = st.columns([1, 2, 1])
with col5:
    uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

# === Info Box If No File Uploaded ===
if not uploaded_file:
    with col2:
        st.info("📁 Please upload a multi-sheet Excel file to get started.")
    st.stop()

try:
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    sheet_names = list(all_sheets.keys())

    # === Search Functionality ===
    if search_query:
        found_in = []
        for sheet_name, df in all_sheets.items():
            if df.empty or df.shape[1] == 0:
                continue
            search_col = df.columns[0]
            values = df[search_col].dropna().apply(
                lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
            )
            if search_query.strip() in values.values:
                found_in.append(sheet_name)

        if found_in:
            st.success(f"✅ '{search_query}' found in: {', '.join(found_in)}")
        else:
            st.warning(f"❌ '{search_query}' not found in any sheet")

    st.divider()

    # === Sidebar Selection ===
    st.sidebar.header("🔧 Sheet Comparison")
    test_sheet = st.sidebar.selectbox("🧪 Test Sheet (shortlist)", sheet_names)
    main_sheet = st.sidebar.selectbox("📘 Main Sheet (with full details)", [s for s in sheet_names if s != test_sheet])

    # === Comparison Logic ===
    if st.sidebar.button("🔍 Compare Now"):
        test_df = all_sheets[test_sheet].copy()
        main_df = all_sheets[main_sheet].copy()

        if 'EMIS No' not in test_df.columns or 'EMIS No' not in main_df.columns:
            st.error("❌ Both sheets must contain an 'EMIS No' column.")
        else:
            # Standardize EMIS No
            test_df['EMIS No'] = test_df['EMIS No'].dropna().apply(
                lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
            )
            main_df['EMIS No'] = main_df['EMIS No'].dropna().apply(
                lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
            )

            # Merge Test with Main on EMIS No
            merged_df = test_df[['EMIS No']].merge(main_df, on='EMIS No', how='left')

            # Add Overlap Status
            merged_df['Overlap Status'] = merged_df['Student Name'].apply(
                lambda x: 'Overlapped' if pd.notna(x) else 'Not Found'
            )

            # Sort: Overlapped first
            merged_df.sort_values(by='Overlap Status', ascending=True, inplace=True)
            merged_df.index = range(1, len(merged_df) + 1)

            # Show Output
            st.success(f"✅ Compared '{test_sheet}' against '{main_sheet}'")
            st.dataframe(merged_df, use_container_width=True)

            # === Excel Export ===
            output = BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Overlap Result"

            for r in dataframe_to_rows(merged_df.reset_index(), index=False, header=True):
                ws.append(r)

            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_length + 2

            wb.save(output)

            st.download_button(
                "📥 Download Overlap Result",
                data=output.getvalue(),
                file_name=f"{test_sheet}_vs_{main_sheet}_overlap.xlsx"
            )

except Exception as e:
    st.error(f"❌ Error processing file: {e}")
