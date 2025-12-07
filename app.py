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

DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)

# ----------------------------
# Helper to read sheets
# ----------------------------
def read_excel(file):
    excel_data = {}
    wb = load_workbook(file)
    for sheet in wb.sheetnames:
        df = pd.read_excel(file, sheet_name=sheet)
        excel_data[sheet] = df
    return excel_data

# ----------------------------
# Email Sender Function
# ----------------------------
def send_email(receiver, subject, body):
    SENDER_EMAIL = st.session_state.get("sender_email", "")
    SENDER_PASS = st.session_state.get("sender_pass", "")

    if not SENDER_EMAIL or not SENDER_PASS:
        return "❌ Email Credentials Missing!"

    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASS)
        server.sendmail(SENDER_EMAIL, receiver, msg.as_string())
        server.quit()
        return "✅ Email Sent Successfully!"
    except Exception as e:
        return f"❌ Error: {str(e)}"

# ----------------------------
# Email Setup Sidebar
# ----------------------------
st.sidebar.header("📧 Email Setup")
st.session_state.sender_email = st.sidebar.text_input("Sender Email")
st.session_state.sender_pass = st.sidebar.text_input("Email App Password", type="password")

# ----------------------------
# Upload Excel
# ----------------------------
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])
if uploaded_file:
    data = read_excel(uploaded_file)

    # List Sheets
    sheet = st.selectbox("📄 Select Sheet", list(data.keys()))
    df = data[sheet]
    st.dataframe(df)

    # Form
    st.subheader("📝 Fill Form Based on Excel Columns")
    form_id = str(uuid.uuid4())
    responses = {}

    with st.form("user_form"):
        for col in df.columns:
            responses[col] = st.text_input(f"{col}")
        submit = st.form_submit_button("Save Response")

        if submit:
            save_path = os.path.join(DATA_DIR, f"{form_id}.json")
            with open(save_path, "w") as f:
                json.dump(responses, f, indent=4)

            st.success("✅ Response Saved!")
            st.json(responses)

            # --------------- Email Send Option ----------------
            st.write("📧 Send Email")
            receiver = st.text_input("To Email:")
            if st.button("Send Email Now"):
                email_text = "\n".join([f"{k}: {v}" for k, v in responses.items()])
                result = send_email(receiver, "Form Submission Result", email_text)
                st.info(result)

# ----------------------------
# Dashboard View
# ----------------------------
st.subheader("📊 Dashboard - Stored Records")
files = [f for f in os.listdir(DATA_DIR) if f.endswith(".json")]

if files:
    all_data = []
    for f in files:
        with open(os.path.join(DATA_DIR, f), "r") as file:
            entry = json.load(file)
            all_data.append(entry)

    df_dashboard = pd.DataFrame(all_data)
    st.dataframe(df_dashboard)
else:
    st.info("No records found!")
