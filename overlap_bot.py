import streamlit as st
import pandas as pd
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# === Streamlit Setup ===
st.set_page_config(page_title="TN Model Schools Overlap Bot", layout="wide")
st.markdown("<h1 style='text-align: center;'>TN Model Schools Student Overlap</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: gray;'>MS CG Team</h4>", unsafe_allow_html=True)
st.divider()

# === Google Sheet Link Input ===
sheet_url = st.text_input("📎 Paste your Google Sheet URL:")

if sheet_url:
    try:
        # Setup API Credentials
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
        client = gspread.authorize(creds)

        # Extract Sheet ID from URL
        if "/d/" in sheet_url:
            sheet_id = sheet_url.split("/d/")[1].split("/")[0]
        else:
            st.error("❌ Invalid Google Sheet URL")
            st.stop()

        # Open Google Sheet and get all tabs (worksheets)
        sheet = client.open_by_key(sheet_id)
        worksheets = sheet.worksheets()

        all_sheets = {}
        for ws in worksheets:
            df = pd.DataFrame(ws.get_all_records())
            all_sheets[ws.title] = df
            time.sleep(1)  # 🕒 Delay to avoid quota limit (429 error)

        sheet_names = list(all_sheets.keys())

        # === Sidebar Options ===
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

                main_df[main_col] = main_df[main_col].apply(
                    lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
                )

                all_compare_values = set()
                for sheet in compare_sheets:
                    comp_df = all_sheets.get(sheet, pd.DataFrame())
                    if not comp_df.empty and comp_df.shape[1] > 0:
                        comp_col = comp_df.columns[0]
                        formatted = comp_df[comp_col].dropna().apply(
                            lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
                        )
                        all_compare_values.update(formatted)
                        time.sleep(0.5)  # optional extra delay if looping many sheets

                main_df["Overlap Status"] = main_df[main_col].isin(all_compare_values).map({
                    True: "Overlapped",
                    False: "Unique"
                })

                main_df.index = range(1, len(main_df) + 1)

                st.success(f"✅ Compared '{main_sheet}' with: {', '.join(compare_sheets)}")
                st.dataframe(main_df, use_container_width=True)

                output = BytesIO()
                main_df.to_excel(output, index=True)
                st.download_button("📥 Download Overlap Result", data=output.getvalue(),
                                   file_name=f"{main_sheet}_vs_overlap.xlsx")

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
                time.sleep(0.2)  # gentle delay for search

            if found_in:
                st.success(f"✅ '{search_query}' found in: {', '.join(found_in)}")
            else:
                st.warning(f"❌ '{search_query}' not found in any sheet")

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("📎 Paste your Google Sheet link above to start.")
