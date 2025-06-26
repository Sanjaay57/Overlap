import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Overlap Checker (Multiple Tabs)", layout="wide")
st.title("📘 Compare One Sheet vs Many")

st.write("Upload an Excel file with multiple tabs. Then compare one sheet against several others (e.g., MSE vs APU, IMU BBA, IITTM...).")

uploaded_file = st.file_uploader("📤 Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        sheet_names = list(all_sheets.keys())

        st.sidebar.header("🔧 Choose Sheets to Compare")
        main_sheet = st.sidebar.selectbox("🧩 Sheet to Check (e.g., MSE)", sheet_names)
        compare_sheets = st.sidebar.multiselect("📌 Compare Against These Sheets", [s for s in sheet_names if s != main_sheet])

        if st.sidebar.button("🔍 Compare"):
            main_df = all_sheets[main_sheet]
            main_col = main_df.columns[0]
            all_compare_values = set()

            for sheet in compare_sheets:
                comp_df = all_sheets[sheet]
                comp_col = comp_df.columns[0]
                all_compare_values.update(comp_df[comp_col].dropna().astype(str).str.strip())

            main_df = main_df.copy()
            main_df[main_col] = main_df[main_col].astype(str).str.strip()
            main_df["Overlap Status"] = main_df[main_col].isin(all_compare_values).map({True: "Selected", False: "Not Selected"})

            st.success(f"✅ Compared '{main_sheet}' with {', '.join(compare_sheets)}")
            st.dataframe(main_df, use_container_width=True)

            # Download result
            output = BytesIO()
            main_df.to_excel(output, index=False)
            st.download_button("📥 Download Result", output.getvalue(), file_name=f"{main_sheet}_vs_multiple_Overlap.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file with multiple tabs to begin.")
