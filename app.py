import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ----------------------------
# CONFIG
# ----------------------------
st.set_page_config(page_title="Excel Form SaaS", layout="wide")
st.title("📄 Excel → Form + Email System")

DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)

META_PATH = os.path.join(DATA_DIR, "meta.json")
RESP_PATH = os.path.join(DATA_DIR, "responses.xlsx")

# ----------------------------
# META
# ----------------------------
def load_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"forms": {}}

def save_meta(meta):
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)

meta = load_meta()

# ----------------------------
# RESPONSES
# ----------------------------
def load_responses():
    if os.path.exists(RESP_PATH):
        return pd.read_excel(RESP_PATH)
    return pd.DataFrame()

def save_responses(df):
    df.to_excel(RESP_PATH, index=False)

# ----------------------------
# FIXED EMAIL FUNCTION (IMPORTANT)
# ----------------------------
def send_email_smtp(sender, password, to_email, subject, message):
    try:
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(message, "plain"))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()
        server.ehlo()

        server.login(sender, password)
        server.send_message(msg)
        server.quit()

        return True
    except Exception as e:
        st.error(f"Email failed for {to_email}: {e}")
        return False

# ----------------------------
# QUERY PARAMS
# ----------------------------
params = st.query_params
mode = params.get("mode", "admin")
form_id = params.get("form_id", None)

# ----------------------------
# FORM VIEW
# ----------------------------
if mode == "form":

    if not form_id or form_id not in meta["forms"]:
        st.error("Invalid form link")
        st.stop()

    form = meta["forms"][form_id]
    st.header(form["form_name"])

    session_id = st.session_state.get("sid", str(uuid.uuid4())[:8])
    st.session_state["sid"] = session_id

    values = {}

    for col in form["columns"]:
        values[col] = st.text_input(col, key=f"{col}_{session_id}")

    if st.button("Submit"):
        df = load_responses()

        row = {
            "FormID": form_id,
            "FormName": form["form_name"],
            "Session": session_id,
            "Time": str(datetime.now())
        }
        row.update(values)

        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        save_responses(df)

        st.success("Submitted!")
        st.rerun()

# ----------------------------
# ADMIN VIEW
# ----------------------------
else:

    st.header("Admin Panel")

    col1, col2 = st.columns(2)

    with col1:
        member_file = st.file_uploader("Member File (Email column required)", type=["xlsx"])
    with col2:
        form_file = st.file_uploader("Form Excel File", type=["xlsx"])

    if member_file and form_file:

        df_members = pd.read_excel(member_file)

        if "Email" not in df_members.columns:
            st.error("Email column missing")
            st.stop()

        df_form = pd.read_excel(form_file)
        df_form.columns = [str(c).strip() for c in df_form.columns]

        st.subheader("Form Preview")
        df_form = st.data_editor(df_form, num_rows="dynamic")

        form_name = st.text_input("Form Name")
        base_url = st.text_input("App URL")
        sender = st.text_input("Gmail")
        password = st.text_input("App Password", type="password")

        # ----------------------------
        # CREATE FORM + SEND EMAIL
        # ----------------------------
        if st.button("Create Form & Send Emails"):

            form_id_new = str(uuid.uuid4())[:8]

            meta["forms"][form_id_new] = {
                "form_name": form_name,
                "columns": list(df_form.columns)
            }
            save_meta(meta)

            link = f"{base_url}?mode=form&form_id={form_id_new}"

            emails = df_members["Email"].dropna().astype(str).tolist()

            success = 0

            for email in emails:
                ok = send_email_smtp(
                    sender,
                    password,
                    email,
                    "Form Invitation",
                    f"Please fill this form:\n{link}"
                )
                if ok:
                    success += 1

            st.success(f"Emails Sent: {success}/{len(emails)}")

    # ----------------------------
    # RESPONSES
    # ----------------------------
    st.markdown("---")
    st.subheader("Responses")

    df = load_responses()

    if not df.empty:
        st.dataframe(df)

        idx = st.selectbox("Select Row", df.index)

        if st.button("Delete"):
            df = df.drop(idx)
            save_responses(df)
            st.rerun()

        if st.button("Save"):
            save_responses(df)
            st.success("Updated")
