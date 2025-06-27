import streamlit as st
import pandas as pd
from io import BytesIO

# === UI Setup ===
st.set_page_config(page_title="TN Model Schools Overlap Bot", layout="wide")
st.markdown("<h1 style='text-align: center;'>TN Model Schools Student Overlap</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>MS CG Team</h4>", unsafe_allow_html=True)
st.divider()

uploaded_file = st.file_uploader("📤 Upload Excel File (.xlsx) with Multiple Sheets", type=["xlsx"])

if uploaded_file:
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        sheet_names = list(all_sheets.keys())

        st.sidebar.header("🔧 Sheet Comparison")
        main_sheet = st.sidebar.selectbox("🧩 Sheet to Check (e.g., MSE)", sheet_names)

        available_compare_sheets = [s for s in sheet_names if s != main_sheet]

        mode = st.sidebar.selectbox("🔽 Select Compare Mode", ["Select sheets manually", "Compare with All"])

        if mode == "Select sheets manually":
            selected_sheets = st.sidebar.multiselect(
                "📌 Compare Against Sheets",
                available_compare_sheets,
                default=[]
            )
        else:
            selected_sheets = available_compare_sheets  # All sheets except main

        if st.sidebar.button("🔍 Compare Now"):
            main_df = all_sheets[main_sheet].copy()

            if main_df.empty or main_df.shape[1] < 2:
                st.error("❌ The main sheet must have at least 2 columns (EMIS and Name).")
            else:
                emis_col = main_df.columns[0]
                name_col = main_df.columns[2] if main_df.shape[1] > 2 else main_df.columns[1]

                main_df[emis_col] = main_df[emis_col].apply(
                    lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
                )

                result_df = main_df[[emis_col, name_col]].copy()
                match_found = []

                for sheet in selected_sheets:
                    comp_df = all_sheets.get(sheet, pd.DataFrame())
                    if not comp_df.empty and comp_df.shape[1] > 0:
                        comp_emis_col = comp_df.columns[0]
                        comp_df[comp_emis_col] = comp_df[comp_emis_col].dropna().apply(
                            lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
                        )
                        compare_values = set(comp_df[comp_emis_col].values)
                        result_df[sheet] = result_df[emis_col].apply(lambda x: x if x in compare_values else None)
                        match_found.append(result_df[sheet].notna())

                if match_found:
                    result_df["Overlap Status"] = pd.concat(match_found, axis=1).any(axis=1).map({
                        True: "Overlapped", False: "Unique"
                    })
                else:
                    result_df["Overlap Status"] = "Unique"

                result_df.index = range(1, len(result_df) + 1)

                st.success(f"✅ Compared '{main_sheet}' with: {', '.join(selected_sheets)}")
                st.dataframe(result_df, use_container_width=True)

                output = BytesIO()
                result_df.to_excel(output, index=True)
                st.download_button(
                    "📥 Download Overlap Result",
                    data=output.getvalue(),
                    file_name=f"{main_sheet}_vs_overlap.xlsx"
                )

        # === Search Feature ===
        st.divider()
        st.subheader("🔎 Search Student Across All Sheets")
        search_query = st.text_input("Enter EMIS number or Name")

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

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
else:
    st.info("📁 Please upload a multi-sheet Excel file to get started.")
