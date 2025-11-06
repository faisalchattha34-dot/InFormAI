import streamlit as st
import pandas as pd
import socket
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import os

# ---------- CONFIG ----------
st.set_page_config(page_title="Excel Smart Form", page_icon="📄", layout="centered")
DATA_FILE = "submissions.xlsx"

# ---------- UTILITIES ----------
def get_local_ip():
    """Get the LAN IP address of this machine."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "localhost"
    finally:
        s.close()
    return ip

def save_submission(new_row):
    """Save a single new submission to persistent Excel file."""
    if os.path.exists(DATA_FILE):
        existing = pd.read_excel(DATA_FILE)
        updated = pd.concat([existing, pd.DataFrame([new_row])], ignore_index=True)
    else:
        updated = pd.DataFrame([new_row])
    updated.to_excel(DATA_FILE, index=False)

def load_submissions():
    """Load existing submissions."""
    if os.path.exists(DATA_FILE):
        return pd.read_excel(DATA_FILE)
    return pd.DataFrame()

# ---------- DETECT DROPDOWNS ----------
def detect_dropdowns(uploaded, df):
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
    return dropdown_dict

# ---------- MODE SELECTION ----------
mode = st.query_params.get("mode", ["admin"])[0]  # default admin
ip = get_local_ip()
base_url = f"http://{ip}:8501"
form_link = f"{base_url}/?mode=form"
admin_link = f"{base_url}/?mode=admin"

st.sidebar.markdown(f"📡 **Shareable Links:**")
st.sidebar.markdown(f"- 🧾 [Form Link]({form_link})")
st.sidebar.markdown(f"- 🧑‍💼 [Admin Link]({admin_link})")

# ---------- ADMIN PANEL ----------
if mode == "admin":
    st.title("🧑‍💼 Admin Panel")

    members_file = st.file_uploader("📋 Upload Members Excel (must contain 'Name' and 'Whatsapp')", type=["xlsx"])
    form_file = st.file_uploader("📄 Upload Form Layout Excel", type=["xlsx"])

    if members_file and form_file:
        # Read members
        members_df = pd.read_excel(members_file)
        if not {"Name", "Whatsapp"}.issubset(members_df.columns):
            st.error("❌ Members file must contain 'Name' and 'Whatsapp' columns.")
            st.stop()

        # Prepare form structure
        df_raw = pd.read_excel(form_file, header=None)
        header_row_index = None
        for i in range(len(df_raw)):
            row = df_raw.iloc[i]
            if row.notna().sum() > 2:
                header_row_index = i
                break

        df = pd.read_excel(form_file, header=header_row_index)
        df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]

        dropdown_dict = detect_dropdowns(form_file, df)
        submissions = load_submissions()

        st.success("✅ Form Detected Successfully!")
        st.write("**Detected Columns:**", list(df.columns))
        st.info(f"🔗 Share this link with members: {form_link}")

        # Progress tracking
        if not submissions.empty:
            filled_names = submissions.get("Name", []).tolist()
            pending_df = members_df[~members_df["Name"].isin(filled_names)]
            st.metric("✅ Forms Filled", len(filled_names))
            st.metric("🕓 Pending Forms", len(pending_df))
            st.subheader("📋 Pending Members")
            st.dataframe(pending_df)
        else:
            st.warning("No submissions yet!")

        st.subheader("📥 All Submissions")
        if not submissions.empty:
            st.dataframe(submissions)
            output = BytesIO()
            submissions.to_excel(output, index=False)
            output.seek(0)
            st.download_button(
                "⬇️ Download All Submissions",
                data=output,
                file_name="all_submissions.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("No data submitted yet.")

# ---------- FORM MODE ----------
elif mode == "form":
    st.title("🧾 Fill the Form")

    uploaded = st.file_uploader("📂 Upload the same Form Excel (for structure)", type=["xlsx"])
    if uploaded:
        df_raw = pd.read_excel(uploaded, header=None)
        header_row_index = None
        for i in range(len(df_raw)):
            row = df_raw.iloc[i]
            if row.notna().sum() > 2:
                header_row_index = i
                break

        df = pd.read_excel(uploaded, header=header_row_index)
        df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]
        dropdown_dict = detect_dropdowns(uploaded, df)

        st.success("✅ Form Loaded. Fill below:")
        data = {}
        for col in df.columns:
            if col in dropdown_dict:
                data[col] = st.selectbox(f"{col}", dropdown_dict[col], key=col)
            else:
                data[col] = st.text_input(f"{col}", key=col)

        if st.button("✅ Submit Form"):
            save_submission(data)
            st.success("🎉 Form submitted successfully!")

    else:
        st.info("📥 Please upload the same Excel layout used by the admin.")

else:
    st.error("Invalid mode! Use ?mode=form or ?mode=admin in URL.")
