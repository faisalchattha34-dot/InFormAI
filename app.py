import streamlit as st
import pandas as pd
import os, json, uuid, re, base64
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

# ----------------------------
# Setup
# ----------------------------
st.set_page_config(page_title="📄 Excel → Web Form + Gmail API", layout="centered")
st.title("📄 Excel → Web Form + Auto Email (Gmail API)")

DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")
ALL_RESPONSES_PATH = os.path.join(DATA_DIR, "all_responses.xlsx")

CLIENT_SECRET_FILE = "client_secret.json"
TOKEN_FILE = "token.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# ----------------------------
# Helper functions
# ----------------------------
def load_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_meta(meta):
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

def detect_dropdowns(excel_file, df_columns):
    excel_file.seek(0)
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active
    dropdowns = {}
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            try:
                if dv.type == "list" and dv.formula1:
                    formula = str(dv.formula1).strip('"')
                    if "," in formula:
                        options = [x.strip() for x in formula.split(",")]
                    else:
                        options = []
                    for cell_range in dv.cells:
                        cidx = cell_range.min_col - 1
                        if 0 <= cidx < len(df_columns):
                            dropdowns[df_columns[cidx]] = options
            except Exception:
                continue
    return dropdowns

def get_gmail_service():
    creds = None
    if "credentials" in st.session_state:
        creds = Credentials.from_authorized_user_info(st.session_state["credentials"], SCOPES)
    elif os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        flow = Flow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES, redirect_uri="urn:ietf:wg:oauth:2.0:oob")
        auth_url, _ = flow.authorization_url(prompt="consent")
        st.info("🔗 Click the link below to authorize Gmail access:")
        st.markdown(f"[Authorize Gmail Access]({auth_url})")
        code = st.text_input("Enter the authorization code:")
        if st.button("✅ Confirm Authorization") and code:
            flow.fetch_token(code=code)
            creds = flow.credentials
            st.session_state["credentials"] = json.loads(creds.to_json())
            with open(TOKEN_FILE, "w") as f:
                f.write(creds.to_json())
            st.success("✅ Gmail authorized successfully!")
            st.experimental_rerun()
        return None
    return build("gmail", "v1", credentials=creds)

def send_gmail_message(service, sender, to, subject, body):
    try:
        message = f"From: {sender}\r\nTo: {to}\r\nSubject: {subject}\r\n\r\n{body}"
        encoded = base64.urlsafe_b64encode(message.encode("utf-8")).decode("utf-8")
        send = service.users().messages().send(userId="me", body={"raw": encoded}).execute()
        return True
    except Exception as e:
        st.error(f"❌ Failed to send to {to}: {e}")
        return False

# ----------------------------
# URL Params
# ----------------------------
params = st.experimental_get_query_params()
mode = params.get("mode", ["admin"])[0]
form_id = params.get("form_id", [None])[0]

meta = load_meta()

# ----------------------------
# FORM VIEW
# ----------------------------
if mode == "form":
    if not form_id or "forms" not in meta or form_id not in meta["forms"]:
        st.error("Invalid or missing form ID. Please contact the admin.")
    else:
        info = meta["forms"][form_id]
        st.header(f"🧾 {info['form_name']}")
        if "session_id" not in st.session_state:
            st.session_state["session_id"] = str(uuid.uuid4())[:8]
        session_id = st.session_state["session_id"]
        st.caption(f"🆔 Session ID: {session_id}")
        dropdowns = info.get("dropdowns", {})
        columns = info["columns"]

        with st.form("user_form"):
            values = {}
            for col in columns:
                if col in dropdowns:
                    values[col] = st.selectbox(col, dropdowns[col], key=f"{col}_{session_id}")
                else:
                    values[col] = st.text_input(col, key=f"{col}_{session_id}")
            submitted = st.form_submit_button("✅ Submit")

        if submitted:
            row = {
                "FormID": form_id,
                "FormName": info["form_name"],
                "UserSession": session_id,
                "SubmittedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            row.update(values)
            if not os.path.exists(ALL_RESPONSES_PATH):
                pd.DataFrame(columns=list(row.keys())).to_excel(ALL_RESPONSES_PATH, index=False)
            existing = pd.read_excel(ALL_RESPONSES_PATH)
            for c in row.keys():
                if c not in existing.columns:
                    existing[c] = None
            combined = pd.concat([existing, pd.DataFrame([row])], ignore_index=True)
            combined.to_excel(ALL_RESPONSES_PATH, index=False)
            st.success("🎉 Response saved successfully!")
            st.balloons()

# ----------------------------
# ADMIN VIEW
# ----------------------------
else:
    st.header("🧑‍💼 Admin Panel")
    st.write("Upload two Excel files — Members & Form Source")

    c1, c2 = st.columns(2)
    with c1:
        member_file = st.file_uploader("📋 Upload Member List (must have 'Email')", type=["xlsx"])
    with c2:
        form_file = st.file_uploader("📄 Upload Form Source", type=["xlsx"])

    if member_file and form_file:
        try:
            df_members = pd.read_excel(member_file)
            df_form = pd.read_excel(form_file)
            if "Email" not in df_members.columns:
                st.error("❌ Member file must contain 'Email' column.")
            else:
                df_form.columns = [str(c).strip().replace("_", " ").title() for c in df_form.columns if pd.notna(c)]
                st.success(f"✅ Form fields detected: {len(df_form.columns)}")
                st.write(df_form.columns.tolist())
                dropdowns = detect_dropdowns(form_file, list(df_form.columns))
                if dropdowns:
                    st.info("Detected dropdowns:")
                    st.table(pd.DataFrame([{"Field": k, "Options": ", ".join(v)} for k, v in dropdowns.items()]))
                form_name = st.text_input("Form name:", f"My Form {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                base_url = st.text_input("Your Streamlit public URL (e.g. https://yourapp.streamlit.app)")

                service = get_gmail_service()
                if service and st.button("🚀 Create Form & Send Emails"):
                    form_id = str(uuid.uuid4())[:10]
                    forms = meta.get("forms", {})
                    forms[form_id] = {
                        "form_name": form_name,
                        "columns": list(df_form.columns),
                        "dropdowns": dropdowns,
                        "created_at": datetime.now().isoformat(),
                    }
                    meta["forms"] = forms
                    save_meta(meta)

                    link = f"{base_url.rstrip('/')}/?mode=form&form_id={form_id}"
                    st.success(f"✅ Form created!\n{link}")

                    emails = df_members["Email"].dropna().unique().tolist()
                    subject = f"Form Invitation: {form_name}"
                    message = f"Hello,\n\nPlease fill the form here:\n{link}\n\nThank you!"
                    sender = service.users().getProfile(userId="me").execute()["emailAddress"]

                    sent = 0
                    for e in emails:
                        if send_gmail_message(service, sender, e, subject, message):
                            sent += 1
                    st.success(f"🎉 Successfully sent to {sent} members!")
        except Exception as e:
            st.error(f"Error: {e}")
