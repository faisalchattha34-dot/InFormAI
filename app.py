import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
from io import BytesIO
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ----------------------------
# Setup
# ----------------------------
st.set_page_config(page_title="📄 Excel → Web Form + Auto Email", layout="centered")
st.title("📄 Excel → Web Form + Auto Email Sender")

DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")
ALL_RESPONSES_PATH = os.path.join(DATA_DIR, "all_responses.xlsx")

# ----------------------------
# Helper Functions
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
                        try:
                            rng = formula.split("!")[-1].replace("$", "")
                            a, b = rng.split(":")
                            col_letters = re.match(r"([A-Za-z]+)", a).group(1)
                            start_row = int(re.match(r"[A-Za-z]+([0-9]+)", a).group(1))
                            end_row = int(re.match(r"[A-Za-z]+([0-9]+)", b).group(1))
                            col_idx = column_index_from_string(col_letters)
                            for r in range(start_row, end_row + 1):
                                v = ws.cell(row=r, column=col_idx).value
                                if v is not None:
                                    options.append(str(v))
                        except Exception:
                            options = []
                    for cell_range in dv.cells:
                        try:
                            cidx = cell_range.min_col - 1
                            if 0 <= cidx < len(df_columns):
                                dropdowns[df_columns[cidx]] = options
                        except Exception:
                            continue
            except Exception:
                continue
    return dropdowns

def send_email_to_members(sender_email, password, members, subject, message):
    sent_count = 0
    for email in members:
        try:
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = email
            msg["Subject"] = subject
            msg.attach(MIMEText(message, "plain"))

            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender_email, password)
                server.send_message(msg)

            sent_count += 1
            st.success(f"✅ Email sent to {email}")
        except Exception as e:
            st.error(f"❌ Failed to send to {email}: {e}")
    return sent_count

# ----------------------------
# URL Params
# ----------------------------
params = st.experimental_get_query_params()
mode = params.get("mode", ["admin"])[0]
form_id = params.get("form_id", [None])[0]

meta = load_meta()

# ----------------------------
# FORM VIEW (Users fill form)
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
        st.caption(f"🆔 Your Session ID: {session_id}")

        dropdowns = info.get("dropdowns", {})
        columns = info["columns"]

        with st.form("user_form", clear_on_submit=False):
            values = {}
            for col in columns:
                if col in dropdowns:
                    values[col] = st.selectbox(col, dropdowns[col], key=f"{col}_{session_id}")
                else:
                    values[col] = st.text_input(col, key=f"{col}_{session_id}")

            submitted = st.form_submit_button("✅ Submit Response")

        if submitted:
            row = {
                "FormID": form_id,
                "FormName": info["form_name"],
                "UserSession": session_id,
                "SubmittedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            row.update(values)

            try:
                if not os.path.exists(ALL_RESPONSES_PATH):
                    pd.DataFrame(columns=list(row.keys())).to_excel(ALL_RESPONSES_PATH, index=False)

                existing = pd.read_excel(ALL_RESPONSES_PATH)

                for col in row.keys():
                    if col not in existing.columns:
                        existing[col] = None

                new_row_df = pd.DataFrame([row])
                combined = pd.concat([existing, new_row_df], ignore_index=True)
                combined.to_excel(ALL_RESPONSES_PATH, index=False)

                st.success("🎉 Response saved successfully!")
                st.balloons()
            except Exception as e:
                st.error(f"❌ Error saving data: {e}")

# ----------------------------
# ADMIN VIEW
# ----------------------------
else:
    st.header("🧑‍💼 Admin Panel")
    st.write("Upload two Excel files — Member List & Form Source.")

    col1, col2 = st.columns(2)
    with col1:
        member_file = st.file_uploader("📋 Upload Member List (must have 'Email' column)", type=["xlsx"])
    with col2:
        form_file = st.file_uploader("📄 Upload Form Source File", type=["xlsx"])

    if member_file and form_file:
        try:
            df_members = pd.read_excel(member_file)
            df_form = pd.read_excel(form_file)

            if "Email" not in df_members.columns:
                st.error("❌ Member file must contain an 'Email' column.")
            else:
                df_form.columns = [str(c).strip().replace("_", " ").title() for c in df_form.columns if pd.notna(c)]
                st.success(f"✅ Form fields detected: {len(df_form.columns)}")
                st.write(df_form.columns.tolist())

                dropdowns = detect_dropdowns(form_file, list(df_form.columns))
                if dropdowns:
                    st.info("Detected dropdowns:")
                    st.table(pd.DataFrame([{"Field": k, "Options": ", ".join(v)} for k, v in dropdowns.items()]))

                form_name = st.text_input("Form Name:", value=f"My Form {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                base_url = st.text_input("Your Streamlit App Public URL (e.g. https://yourapp.streamlit.app)")

                sender_email = st.text_input("Your Gmail Address:")
                password = st.text_input("Your Gmail App Password:", type="password")

                if st.button("🚀 Create Form & Send Emails"):
                    if not base_url:
                        st.error("Please enter your app URL to generate the link.")
                    elif not sender_email or not password:
                        st.error("Please enter Gmail and App Password to send emails.")
                    else:
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
                        st.success(f"✅ Form created successfully!\n{link}")

                        st.info("📧 Sending form link to all members...")

                        emails = df_members["Email"].dropna().unique().tolist()
                        subject = f"Form Invitation: {form_name}"
                        message = f"Hello,\n\nPlease fill out the form below:\n{link}\n\nThank you!"

                        sent_count = send_email_to_members(sender_email, password, emails, subject, message)

                        st.success(f"🎉 Email sent to {sent_count} members successfully!")

        except Exception as e:
            st.error(f"❌ Error processing files: {e}")
