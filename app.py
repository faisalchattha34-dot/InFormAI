import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="📄 Excel Form SaaS PRO", layout="wide")
st.title("📄 Excel → Web Form + Email SaaS (FIXED)")

RESP_FILE = "responses.xlsx"
META_FILE = "meta.json"

# =========================
# INIT FILES
# =========================
if not os.path.exists(META_FILE):
    with open(META_FILE, "w") as f:
        json.dump({"forms": {}}, f)

if not os.path.exists(RESP_FILE):
    df_init = pd.DataFrame(columns=["form_id", "response_id", "timestamp", "data"])
    df_init.to_excel(RESP_FILE, index=False)

# =========================
# LOAD META
# =========================
def load_meta():
    with open(META_FILE, "r") as f:
        return json.load(f)

def save_meta(meta):
    with open(META_FILE, "w") as f:
        json.dump(meta, f, indent=4)

meta = load_meta()

# =========================
# LOAD / SAVE RESPONSES
# =========================
def load_responses():
    return pd.read_excel(RESP_FILE)

def save_responses(df):
    df.to_excel(RESP_FILE, index=False)

# =========================
# SMTP EMAIL
# =========================
def send_email(receiver_email, form_link):
    sender_email = "your_email@gmail.com"
    sender_pass = "your_app_password"

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg["Subject"] = "Fill this Form"

    body = f"Please fill this form:\n{form_link}"
    msg.attach(MIMEText(body, "plain"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_pass)
        server.send_message(msg)
        server.quit()
    except Exception as e:
        st.error(f"Email failed: {e}")

# =========================
# CREATE FORM
# =========================
st.sidebar.header("📋 Create Form")

form_name = st.sidebar.text_input("Form Name")
uploaded_file = st.sidebar.file_uploader("Upload Excel (Fields)", type=["xlsx"])

if st.sidebar.button("Create Form"):
    if form_name and uploaded_file:
        df = pd.read_excel(uploaded_file)
        columns = df.columns.tolist()

        form_id = str(uuid.uuid4())

        meta["forms"][form_id] = {
            "name": form_name,
            "columns": columns,
            "created_at": str(datetime.now())
        }
        save_meta(meta)

        st.sidebar.success(f"Form Created: {form_name}")
        st.sidebar.write(columns)

# =========================
# SHARE FORM
# =========================
st.sidebar.header("📧 Send Form")

if meta["forms"]:
    form_options = {v["name"]: k for k, v in meta["forms"].items()}
    selected_name = st.sidebar.selectbox("Select Form", list(form_options.keys()))
    form_id = form_options[selected_name]

    emails = st.sidebar.text_area("Emails (comma separated)")

    if st.sidebar.button("Send"):
        email_list = [e.strip() for e in emails.split(",") if e.strip()]
        link = f"http://localhost:8501/?form_id={form_id}"

        for e in email_list:
            send_email(e, link)

        st.sidebar.success("Emails Sent!")

# =========================
# FORM VIEW
# =========================
query = st.experimental_get_query_params()

if "form_id" in query:
    fid = query["form_id"][0]

    if fid in meta["forms"]:
        form = meta["forms"][fid]

        st.subheader(f"📝 {form['name']}")

        user_data = {}
        for col in form["columns"]:
            user_data[col] = st.text_input(col)

        if st.button("Submit"):
            df = load_responses()

            new_row = {
                "form_id": fid,
                "response_id": str(uuid.uuid4()),
                "timestamp": str(datetime.now()),
                "data": json.dumps(user_data)
            }

            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_responses(df)

            st.success("Submitted Successfully!")
            st.rerun()

# =========================
# ADMIN DASHBOARD
# =========================
st.markdown("---")
st.subheader("📊 Admin Dashboard")

df = load_responses()

if df.empty:
    st.info("No responses yet.")
else:
    for i, row in df.iterrows():
        st.markdown(f"**ID:** {row['response_id']} | {row['timestamp']}")

        data = json.loads(row["data"])

        updated = {}
        cols = st.columns([3, 1, 1])

        with cols[0]:
            for k, v in data.items():
                updated[k] = st.text_input(
                    k,
                    value=v,
                    key=f"{row['response_id']}_{k}"
                )

        with cols[1]:
            if st.button("Save", key=f"save_{row['response_id']}"):
                df.loc[i, "data"] = json.dumps(updated)
                save_responses(df)
                st.success("Updated!")

        with cols[2]:
            if st.button("Delete", key=f"del_{row['response_id']}"):
                df = df[df["response_id"] != row["response_id"]]
                save_responses(df)
                st.success("Deleted!")
                st.rerun()
