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
st.title("📧 Gmail Auto Sender via Excel Sheet")

CLIENT_SECRET_FILE = "client_secret.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
TOKEN_FILE = "token.json"

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
            redirect_uri = "https://informai-nbjyzvkxi8eghb688vcxvu.streamlit.app/"
            flow = Flow.from_client_secrets_file(
                CLIENT_SECRET_FILE, scopes=SCOPES, redirect_uri=redirect_uri
            )
            auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline")

            st.markdown("### 🔐 Gmail Authorization Required")
            st.write("Click below to connect your Gmail account:")
            st.markdown(f"[Authorize Gmail Access]({auth_url})")

            query_params = st.experimental_get_query_params()
            if "code" in query_params:
                code = query_params["code"][0]
                flow.fetch_token(code=code)
                creds = flow.credentials
                with open(TOKEN_FILE, "w") as token:
                    token.write(creds.to_json())
                st.experimental_set_query_params()  # clear URL params
                st.success("✅ Gmail authorized successfully! You can now send emails.")
            else:
                st.stop()
    service = build("gmail", "v1", credentials=creds)
    return service

# ----------------------------
# SEND EMAIL FUNCTION
# ----------------------------
def send_email(service, to, subject, message_text):
    from email.mime.text import MIMEText
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

    st.subheader("📁 Step 1: Upload Form File (for generating form link)")
    form_file = st.file_uploader("Upload Excel file for form (columns extraction)", type=["xlsx"])

    st.subheader("📧 Step 2: Upload Email List File")
    email_file = st.file_uploader("Upload Excel file containing Name and Email columns", type=["xlsx"])

    if form_file and email_file:
        # Read email list
        df_email = pd.read_excel(email_file)

        # Validate columns
        if "Email" not in df_email.columns:
            st.error("The email list file must contain an 'Email' column.")
            st.stop()
        if "Name" not in df_email.columns:
            st.warning("No 'Name' column found — emails will not be personalized.")

        st.success(f"✅ Loaded {len(df_email)} recipients from email list.")
        st.write(df_email.head())

        # Dummy form link — can be dynamic later
        form_link = st.text_input(
            "🔗 Enter your generated form link:",
            "https://informai-nbjyzvkxi8eghb688vcxvu.streamlit.app/?mode=form&form_id=YOUR_ID"
        )

        subject = st.text_input("📨 Email Subject", "Please fill this form")
        body = st.text_area(
            "💬 Email Message (you can use {{name}} and {{link}} placeholders)",
            "Hello {{name}},<br><br>Please fill this form: {{link}}<br><br>Thank you!",
            height=150,
        )

        if st.button("🚀 Send Emails"):
            success = 0
            for _, row in df_email.iterrows():
                try:
                    email = row["Email"]
                    name = row["Name"] if "Name" in df_email.columns else ""
                    message = body.replace("{{link}}", form_link).replace("{{name}}", name)
                    send_email(service, email, subject, message)
                    success += 1
                except Exception as e:
                    st.warning(f"Failed to send to {row['Email']}: {e}")

            st.success(f"✅ Emails sent successfully to {success} recipients!")
else:
    st.error("❌ client_secret.json file not found. Please upload your Google OAuth credentials.")
