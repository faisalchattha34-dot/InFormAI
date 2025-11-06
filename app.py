import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import os

# -------------------------
# PAGE CONFIG
# -------------------------
st.set_page_config(page_title="Excel Smart Form", page_icon="📄", layout="centered")

# -------------------------
# MODE DETECTION
# -------------------------
query_params = st.experimental_get_query_params()
mode = query_params.get("mode", ["admin"])[0]  # default admin mode

# -------------------------
# FILE NAMES
# -------------------------
FORM_TEMPLATE_FILE = "uploaded_form.xlsx"
RESPONSES_FILE = "responses.xlsx"

# -------------------------
# DROPDOWN DETECTION FUNCTION
# -------------------------
def detect_dropdowns(excel_file, df_columns):
    excel_file.seek(0)
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active
    dropdown_dict = {}

    if ws.data_validations is not None:
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1:
                formula = str(dv.formula1).strip('"')
                if "," in formula:
                    values = [v.strip() for v in formula.split(",")]
                else:
                    values = []
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
                        if 0 <= col_index < len(df_columns):
                            dropdown_dict[df_columns[col_index]] = values
                    except Exception:
                        continue
    return dropdown_dict

# -------------------------
# MODE: ADMIN
# -------------------------
if mode == "admin":
    st.title("🧑‍💼 Admin Panel — Form Setup")

    uploaded = st.file_uploader("📂 Upload Excel Form Template", type=["xlsx"])
    if uploaded:
        try:
            df_raw = pd.read_excel(uploaded, header=None)
            header_row_index = None
            for i in range(len(df_raw)):
                row = df_raw.iloc[i]
                if row.notna().sum() > 2:
                    header_row_index = i
                    break

            df = pd.read_excel(uploaded, header=header_row_index)
            df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]

            st.success("✅ Columns Detected Successfully!")
            st.write("**Detected Columns:**", list(df.columns))

            dropdowns = detect_dropdowns(uploaded, list(df.columns))
            if dropdowns:
                st.info("🎯 **Detected Dropdown Columns:**")
                st.table(pd.DataFrame([
                    {"Column": col, "Dropdown Options": ", ".join(vals)}
                    for col, vals in dropdowns.items()
                ]))
            else:
                st.warning("⚠️ No dropdown lists detected.")

            # Save the uploaded Excel form template
            with open(FORM_TEMPLATE_FILE, "wb") as f:
                f.write(uploaded.getbuffer())
            st.success("📄 Form template saved successfully!")

            st.markdown("---")
            st.info("✅ Share this form link with anyone:")
            st.code("http://localhost:8501/?mode=form", language="bash")

            st.caption("ℹ️ Replace localhost with your network IP or deploy link if needed.")

        except Exception as e:
            st.error(f"❌ Error: {e}")

    # Show saved responses
    if os.path.exists(RESPONSES_FILE):
        st.subheader("📊 Submitted Responses (Live Data)")
        df_responses = pd.read_excel(RESPONSES_FILE)
        st.dataframe(df_responses)
        st.download_button(
            "⬇️ Download All Responses",
            data=open(RESPONSES_FILE, "rb").read(),
            file_name="responses.xlsx"
        )
    else:
        st.info("No responses yet.")

# -------------------------
# MODE: FORM
# -------------------------
elif mode == "form":
    st.title("📝 Fill the Form")

    if not os.path.exists(FORM_TEMPLATE_FILE):
        st.error("❌ No form template found. Please ask the admin to upload one.")
    else:
        df = pd.read_excel(FORM_TEMPLATE_FILE)
        df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]
        dropdowns = detect_dropdowns(open(FORM_TEMPLATE_FILE, "rb"), list(df.columns))

        form_data = {}
        for col in df.columns:
            if col in dropdowns:
                form_data[col] = st.selectbox(col, dropdowns[col])
            else:
                form_data[col] = st.text_input(col)

        if st.button("✅ Submit"):
            new_row = pd.DataFrame([form_data])

            if os.path.exists(RESPONSES_FILE):
                old = pd.read_excel(RESPONSES_FILE)
                updated = pd.concat([old, new_row], ignore_index=True)
            else:
                updated = new_row

            updated.to_excel(RESPONSES_FILE, index=False)
            st.success("🎉 Your response has been submitted successfully!")

            st.write("✅ Thank you for filling out the form!")

else:
    st.error("Invalid mode. Use ?mode=admin or ?mode=form in the URL.")
