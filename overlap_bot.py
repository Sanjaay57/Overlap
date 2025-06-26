import streamlit as st
import pandas as pd
from io import BytesIO

# === Page Config ===
st.set_page_config(page_title="TN Model Schools Overlap Bot", layout="wide")

# === App Title ===
st.markdown("<h1 style='text-align: center;'>TN Model Schools Student Overlap</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>MS CG Team</h4>", unsafe_allow_html=True)
st.divider()

# === Upload Section ===
uploaded_file = st.file_uploader("📤 Upload Excel File (.xlsx) with Multiple Sheets", type=["xlsx"])

if uploaded_file:
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        sheet_names = list(all_sheets.keys())

        # === Sidebar Controls ===
        st.sidebar.header("🔧 Sheet Comparison")
        main_sheet = st.sidebar.selectbox("🧩 Sheet to Check (e.g., MSE)", sheet_names)
        compare_sheets = st.sidebar.multiselect(
            "📌 Compare Against These Sheets",
            [s for s in sheet_names if s != main_sheet]
        )

        # === Compare Logic ===
        if st.sidebar.button("🔍 Compare Now"):
            main_df = all_sheets[main_sheet]
            main_col = main_df.columns[0]
            all_compare_values = set()

            for sheet in compare_sheets:
                comp_df = all_sheets[sheet]
                comp_col = comp_df.columns[0]
                all_compare_values.update(comp_df[comp_col].dropna().astype(str).str.strip())

            main_df = main_df.copy()
            main_df[main_col] = main_df[main_col].astype(str).str.strip()
            main_df["Overlap Status"] = main_df[main_col].isin(all_compare_values).map({
                True: "Selected",
                False: "Not Selected"
            })

            st.success(f"✅ Compared **{main_sheet}** with: {', '.join(compare_sheets)}")
            st.dataframe(main_df, use_container_width=True)

            # Excel Download
            output = BytesIO()
            main_df.to_excel(output, index=False)
            st.download_button(
                "📥 Download Overlap Result",
                data=output.getvalue(),
                file_name=f"{main_sheet}_vs_multiple_overlap.xlsx"
            )

        # === Search Section ===
        st.divider()
        st.subheader("🔎 Search Student Across All Sheets")
        search_query = st.text_input("Enter EMIS number or Name")

        if search_query:
            found_in = []
            for sheet_name, df in all_sheets.items():
                values = df[df.columns[0]].astype(str).str.strip()
                if search_query.strip() in values.values:
                    found_in.append(sheet_name)

            if found_in:
                st.success(f"✅ '{search_query}' found in: {', '.join(found_in)}")
            else:
                st.warning(f"❌ '{search_query}' not found in any sheet")

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
else:
    st.info("📁 Please upload a multi-sheet Excel file to get started.")
