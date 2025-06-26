import streamlit as st
import pandas as pd
from io import BytesIO

# === Page Config ===
st.set_page_config(page_title="TN Model Schools Overlap Bot", layout="wide")

# === App Title ===
st.markdown("<h1 style='text-align: center;'>TN Model Schools Student Overlap</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>MS CG Team</h4>", unsafe_allow_html=True)
st.divider()

# === File Upload ===
uploaded_file = st.file_uploader("📤 Upload Excel File (.xlsx) with Multiple Sheets", type=["xlsx"])

if uploaded_file:
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        sheet_names = list(all_sheets.keys())

        # === Sidebar Selections ===
        st.sidebar.header("🔧 Sheet Comparison")
        main_sheet = st.sidebar.selectbox("🧩 Sheet to Check (e.g., MSE)", sheet_names)
        compare_sheets = st.sidebar.multiselect(
            "📌 Compare Against These Sheets",
            [s for s in sheet_names if s != main_sheet]
        )

        if st.sidebar.button("🔍 Compare Now"):
            main_df = all_sheets[main_sheet]

            if main_df.empty or main_df.shape[1] == 0:
                st.error("❌ The main sheet is empty or has no columns.")
            else:
                main_df = main_df.copy()
                main_col = main_df.columns[0]

                # Clean EMIS values in main sheet
                main_df[main_col] = main_df[main_col].apply(
                    lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
                )

                # Combine EMIS from other sheets
                all_compare_values = set()
                for sheet in compare_sheets:
                    comp_df = all_sheets.get(sheet, pd.DataFrame())
                    if not comp_df.empty and comp_df.shape[1] > 0:
                        comp_col = comp_df.columns[0]
                        formatted_values = comp_df[comp_col].dropna().apply(
                            lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
                        )
                        all_compare_values.update(formatted_values)

                # Compare for overlap
                main_df["Overlap Status"] = main_df[main_col].isin(all_compare_values).map({
                    True: "Overlapped",
                    False: "Unique"
                })

                # Insert S.No column (starting from 1)
                main_df.insert(0, "S.No", range(1, len(main_df) + 1))

                st.success(f"✅ Compared '{main_sheet}' with: {', '.join(compare_sheets)}")
                st.dataframe(main_df, use_container_width=True)

                # Download option
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
