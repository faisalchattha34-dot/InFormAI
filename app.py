import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime
from openpyxl import load_workbook
from io import BytesIO
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ----------------------------
# Setup
# ----------------------------
st.set_page_config(page_title="📄 Excel → Web Form + Auto Email", layout="wide")
st.title("📄 Excel → Web Form + Auto Email Sender + Dashboard")

# ----------------------------
# CSS
# ----------------------------
st.markdown("""
<style>
:root { color-scheme: light dark; }
body { font-family: 'Arial'; }
.stButton>button { background:#3498db;color:white;border-radius:8px; }
</style>
""", unsafe_allow_html=True)

# ----------------------------
# Paths
# ----------------------------
DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")
RESP_PATH = os.path.join(DATA_DIR, "all_responses.xlsx")

# ----------------------------
def load_meta():
    return json.load(open(META_PATH, "r")) if os.path.exists(META_PATH) else {}

def save_meta(meta):
    json.dump(meta, open(META_PATH, "w"), indent=2)

def load_responses():
    return pd.read_excel(RESP_PATH) if os.path.exists(RESP_PATH) else pd.DataFrame()

def save_responses(df):
    df.to_excel(RESP_PATH, index=False)

# ----------------------------
# AUTO DROPDOWN DETECTION FROM DATA
# ----------------------------
def detect_dropdowns_from_data(df):
    dropdowns = {}
    for col in df.columns:
        unique_vals = df[col].dropna().unique().tolist()
        if 1 < len(unique_vals) <= 50:
            dropdowns[col] = unique_vals
    return dropdowns

# ----------------------------
# URL Params
# ----------------------------
params = st.experimental_get_query_params()
mode = params.get("mode", ["admin"])[0]
form_id = params.get("form_id", [None])[0]
meta = load_meta()

# ----------------------------
# USER FORM MODE
# ----------------------------
if mode == "form":
    if not form_id or "forms" not in meta or form_id not in meta["forms"]:
        st.error("Invalid Form Link!")
    else:
        info = meta["forms"][form_id]
        st.header(f"📝 {info['form_name']}")
        dropdowns = info.get("dropdowns", {})
        columns = info["columns"]

        with st.form("user_submit"):
            values = {}
            for col in columns:
                if col in dropdowns:
                    options = dropdowns[col] + ["✏ Enter custom value"]
                    selected = st.selectbox(col, options)
                    if selected == "✏ Enter custom value":
                        values[col] = st.text_input(f"Enter new value for '{col}'")
                    else:
                        values[col] = selected
                else:
                    values[col] = st.text_input(col)

            submit = st.form_submit_button("Submit Response")

        if submit:
            row = {"FormID": form_id, "SubmittedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            row.update(values)
            res = load_responses()
            for c in row:
                if c not in res.columns: res[c] = None
            res = pd.concat([res, pd.DataFrame([row])], ignore_index=True)
            save_responses(res)
            st.success("Response Saved Successfully 🎉")

# ----------------------------
# ADMIN PANEL
# ----------------------------
else:
    st.header("🧑‍💼 Admin Panel")

    col1, col2 = st.columns(2)
    with col1:
        member_file = st.file_uploader("Upload Members File (Must have Email)", type=["xlsx"])
    with col2:
        form_file = st.file_uploader("Upload Source Excel", type=["xlsx"])

    if member_file and form_file:
        members = pd.read_excel(member_file)
        df_form = pd.read_excel(form_file)

        st.subheader("Editable Sheet Preview")
        if "current" not in st.session_state:
            st.session_state.current = df_form.copy()

        edited = st.data_editor(st.session_state.current, num_rows="dynamic")
        st.session_state.current = edited

        dropdowns = detect_dropdowns_from_data(st.session_state.current)

        st.write("Detected Dropdown Fields:")
        if dropdowns:
            st.table(pd.DataFrame([{"Field": k, "Options": ",".join(v)} for k,v in dropdowns.items()]))
        else:
            st.info("No dropdowns detected.")

        form_name = st.text_input("Form Name", value="Generated Form")
        app_url = st.text_input("App Public URL (Example: https://myapp.streamlit.app)")
        sender = st.text_input("Gmail Sender")
        password = st.text_input("Gmail App Password", type="password")

        if st.button("Create Form & Send Emails"):
            fid = str(uuid.uuid4())[:10]
            meta.setdefault("forms", {})
            meta["forms"][fid] = {"form_name": form_name, "columns": list(st.session_state.current.columns),
                                  "dropdowns": dropdowns}
            save_meta(meta)

            link = f"{app_url}/?mode=form&form_id={fid}"
            st.success(f"Form Created Successfully\n{link}")

# ----------------------------
# RESPONSE DASHBOARD
# ----------------------------
    st.subheader("📊 Responses")
    res = load_responses()
    if res.empty:
        st.info("No responses yet.")
    else:
        st.dataframe(res)
