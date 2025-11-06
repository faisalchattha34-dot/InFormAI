import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import os

st.set_page_config(page_title="Excel Smart Form", page_icon="📄", layout="centered")

# --- Detect mode (admin or form)
query_params = st.experimental_get_query_params()
mode = query_params.get("mode", ["admin"])[0]

# --- File path for permanent submissions
SUBMISSION_FILE = "submissions.xlsx"

def load_submissions():
    if os.path.exists(SUBMISSION_FILE):
        return pd.read_excel(SUBMISSION_FILE)
    return pd.DataFrame()

def save_submissions(df):
    df.to_excel(SUBMISSION_FILE, index=False)

# -------------------- ADMIN MODE --------------------
if mode == "admin":
    st.title("👩‍💼 Admin Dashboard - Excel Smart Form")

    uploaded = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

    if uploaded:
        try:
            # Step 1: Detect header row
            df_raw = pd.read_excel(uploaded, header=None)
            header_row_index = None
            for i in range(len(df_raw)):
                row = df_raw.iloc[i]
                if row.notna().sum() > 2:
                    header_row_index = i
                    break

            # Step 2: Read DataFrame
            df = pd.read_excel(uploaded, header=header_row_index)
            df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]

            st.success("✅ Columns Detected Successfully!")
            st.write("**Detected Columns:**", list(df.columns))

            # Step 3: Detect dropdowns (data validation)
            uploaded.seek(0)
            wb = load_workbook(uploaded, data_only=True)
            ws = wb.active

            dropdown_dict = {}
            if ws.data_validations is not None:
                for dv in ws.data_validations.dataValidation:
                    if dv.type == "list" and dv.formula1:
                        formula = str(dv.formula1).strip('"')
                        if "," in formula:
                            values = [v.strip() for v in formula.split(",")]

                            for cell_range in dv.cells:
                                try:
                                    if hasattr(cell_range, "min_col"):
                                        col_index = cell_range.min_col - 1
                                    else:
                                        s = str(cell_range).split(":")[0]
                                        match = re.match(r"([A-Za-z]+)", s)
                                        if not match:
                                            continue
                                        col_letters = match.group(1)
                                        col_index = column_index_from_string(col_letters) - 1

                                    if 0 <= col_index < len(df.columns):
                                        dropdown_dict[df.columns[col_index]] = values
                                except Exception:
                                    continue

            # Step 4: Show dropdowns
            if dropdown_dict:
                st.info("🎯 **Detected Dropdown Columns:**")
                st.table(pd.DataFrame([
                    {"Column": col, "Dropdown Options": ", ".join(vals)}
                    for col, vals in dropdown_dict.items()
                ]))
            else:
                st.warning("⚠️ No dropdown lists detected in this Excel file.")

            # Step 5: Show form link
            st.markdown("### 🔗 Shareable Form Link")
            app_url = st.experimental_get_query_params()
            base_url = st.get_option("browser.serverAddress")
            port = st.get_option("browser.serverPort")
            form_link = f"http://{base_url}:{port}/?mode=form"
            st.code(form_link, language="markdown")
            st.info("Send this link to all members via WhatsApp or email — everyone uses the same form link.")

            # Step 6: Show submissions tracking
            st.subheader("📊 Form Submissions Tracker")
            submitted_df = load_submissions()
            if not submitted_df.empty:
                st.success(f"✅ Total Submissions: {len(submitted_df)}")
                st.dataframe(submitted_df)
            else:
                st.warning("No submissions yet.")

            # Step 7: Download all submissions
            if not submitted_df.empty:
                buf = BytesIO()
                submitted_df.to_excel(buf, index=False)
                buf.seek(0)
                st.download_button(
                    label="⬇️ Download All Submissions",
                    data=buf,
                    file_name="submissions.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ Unexpected error: {e}")

# -------------------- FORM MODE --------------------
elif mode == "form":
    st.title("📝 Member Form - Excel Smart Form")

    # Load the structure from the last uploaded Excel (admin must upload at least once)
    if not os.path.exists("submissions_structure.xlsx"):
        st.error("⚠️ Admin must first upload Excel in Admin Mode to set up form structure.")
    else:
        df_structure = pd.read_excel("submissions_structure.xlsx")

        # Load dropdowns (stored separately)
        dropdown_file = "dropdowns.pkl"
        import pickle
        if os.path.exists(dropdown_file):
            with open(dropdown_file, "rb") as f:
                dropdown_dict = pickle.load(f)
        else:
            dropdown_dict = {}

        st.subheader("Please fill the following form:")
        data = {}
        for col in df_structure.columns:
            if col in dropdown_dict:
                data[col] = st.selectbox(col, dropdown_dict[col])
            else:
                data[col] = st.text_input(col)

        if st.button("✅ Submit Form"):
            new_row = pd.DataFrame([data])
            existing = load_submissions()
            updated = pd.concat([existing, new_row], ignore_index=True)
            save_submissions(updated)
            st.success("🎉 Your response has been submitted successfully!")
