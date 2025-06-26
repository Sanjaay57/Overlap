import streamlit as st
import pandas as pd
from io import BytesIO

# === Streamlit Page Config ===
st.set_page_config(page_title="TN Model Schools Overlap Bot", layout="wide")
st.markdown("<h1 style='text-align: center;'>TN Model Schools Student Overlap</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>MS CG Team</h4>", unsafe_allow_html=True)
st.divider()

# === Step 1: Google Sheet Link Input ===
sheet_url = st.text_input("https://docs.google.com/spreadsheets/d/196KjH5zEq8D4_I1OuLcrtYjhUYYZrOfZTm9IOfrtqX8/edit?usp=sharing")

# === Step 2: Process Google Sheet ===
all_sheets = {}
if sheet_url:
    try:
        export_url = sheet_url.replace("/edit#gid=", "/export?format=xlsx&gid=")
        all_sheets = pd.read_excel(export_url, sheet_name=None)
        st.success("✅ Google Sheet loaded successfully.")
    except Exception as e:
        st.error(f"❌ Failed to load from Google Sheet: {e}")

# === Continue if Google Sheet is Loaded ===
if all_sheets:
    sheet_names = list(all_sheets.keys())

    # === Sidebar UI ===
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

            # Format main values (remove .000000)
            main_df[main_col] = main_df[main_col].apply(
                lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
            )

            # Collect all EMIS/IDs from compare sheets
            all_compare_values = set()
            for sheet in compare_sheets:
                comp_df = all_sheets.get(sheet, pd.DataFrame())
                if not comp_df.empty and comp_df.shape[1] > 0:
                    comp_col = comp_df.columns[0]
                    formatted_values = comp_df[comp_col].dropna().apply(
                        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
                    )
                    all_compare_values.update(formatted_values)

            # Mark overlap
            main_df["Overlap Status"] = main_df[main_col].isin(all_compare_values).map({
                True: "Overlapped",
                False: "Unique"
            })

            # Start index from 1
            main_df.index = range(1, len(main_df) + 1)

            st.success(f"✅ Compared '{main_sheet}' with: {', '.join(compare_sheets)}")
            st.dataframe(main_df, use_container_width=True)

            # Download button
            output = BytesIO()
            main_df.to_excel(output, index=True)
            st.download_button(
                "📥 Download Overlap Result",
                data=output.getvalue(),
                file_name=f"{main_sheet}_vs_multiple_overlap.xlsx"
            )

    # === Search Bar ===
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
else:
    st.info("📎 Paste a public Google Sheet link above to start.")
