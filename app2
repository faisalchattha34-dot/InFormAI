import streamlit as st
import pandas as pd
import os
import base64
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# ----------------------------
# CONFIG
# ----------------------------
st.set_page_config(page_title="📧 Gmail Auto-Sender", layout="centered")

CLIENT_SECRET_FILE = "client_secret.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
TOKEN_FILE = "token.json"

st.title("📧 Gmail Auto Sender via Excel Sheet")

# ----------------------------
# AUTH SECTION
# ----------------------------
def get_gmail_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = Flow.from_client_secrets_file(
                CLIENT_SECRET_FILE,
                scopes=SCOPES,
                redirect_uri="https://informai-nbjyzvkxi8eghb688vcxvu.streamlit.app/",
            )
            auth_url, _ = flow.authorization_url(prompt="consent")
            st.markdown(f"[🔐 Click here to authorize Gmail access]({auth_url})")
            st.stop()
    service = build("gmail", "v1", credentials=creds)
    return service

# ----------------------------
# SEND EMAIL FUNCTION
# ----------------------------
def send_email(service, to, subject, message_text):
    from email.mime.text import MIMEText
    import base64

    message = MIMEText(message_text, "html")
    message["to"] = to
    message["subject"] = subject

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    body = {"raw": raw}
    service.users().messages().send(userId="me", body=body).execute()

# ----------------------------
# APP LOGIC
# ----------------------------
if os.path.exists(CLIENT_SECRET_FILE):
    st.info("✅ Gmail API credentials loaded.")
    service = get_gmail_service()

    uploaded = st.file_uploader("📂 Upload Excel File (must include 'Email' column)", type=["xlsx"])

    if uploaded:
        df = pd.read_excel(uploaded)
        if "Email" not in df.columns:
            st.error("Excel file must have an 'Email' column.")
        else:
            st.success(f"Found {len(df)} emails.")
            st.write(df.head())

            form_link = st.text_input("🔗 Enter your form link to send:",
                                      "https://informai-nbjyzvkxi8eghb688vcxvu.streamlit.app/?mode=form&form_id=YOUR_ID")

            subject = st.text_input("📨 Email Subject", "Please fill this form")
            body = st.text_area(
                "💬 Email Message (you can use {{link}} placeholder)",
                "Hello! Kindly fill this form: {{link}}<br><br>Thank you!",
                height=150,
            )

            if st.button("🚀 Send Emails"):
                success = 0
                for idx, row in df.iterrows():
                    try:
                        email = row["Email"]
                        message = body.replace("{{link}}", form_link)
                        send_email(service, email, subject, message)
                        success += 1
                    except Exception as e:
                        st.warning(f"Failed to send to {row['Email']}: {e}")

                st.success(f"✅ Emails sent successfully to {success} recipients!")
else:
    st.error("❌ client_secret.json file not found. Please upload your Google OAuth credentials.")
