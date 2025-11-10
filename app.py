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
st.set_page_config(page_title="📄 Excel → Web Form + Auto Email", layout="wide")
st.title("📄 Excel → Web Form + Auto Email Sender + Dashboard")

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
                    for cell_range in dv.cells:
                        cidx = cell_range.min_col - 1
                        if 0 <= cidx < len(df_columns):
                            dropdowns[df_columns[cidx]] = options
            except Exception:
                continue
    return dropdowns

def send_email_to_members(sender_email, password, members, subject, message):
    sent_count = 0
    results = []
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
            results.append({"Email": email, "Status": "✅ Sent"})
        except Exception as e:
            results.append({"Email": email, "Status": f"❌ Failed ({e})"})
    return sent_count, results

def load_responses():
    if os.path.exists(ALL_RESPONSES_PATH):
        return pd.read_excel(ALL_RESPONSES_PATH)
    return pd.DataFrame()

def save_responses(df):
    df.to_excel(ALL_RESPONSES_PATH, index=False)

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
                responses = load_responses()
                for col in row.keys():
                    if col not in responses.columns:
                        responses[col] = None

                responses = pd.concat([responses, pd.DataFrame([row])], ignore_index=True)
                save_responses(responses)

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
                base_url = st.text_input("Your Streamlit App Public URL (example: https://yourapp.streamlit.app)")

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

                        sent_count, send_results = send_email_to_members(sender_email, password, emails, subject, message)

                        st.success(f"🎉 Emails sent: {sent_count}/{len(emails)}")

                        st.subheader("📧 Email Send Status")
                        st.table(pd.DataFrame(send_results))

    st.markdown("---")
    st.subheader("📊 Responses Dashboard")

    responses = load_responses()
    if responses.empty:
        st.info("No responses submitted yet.")
    else:
        form_filter = st.selectbox("Select Form to View Responses:", ["All"] + [f["form_name"] for f in meta.get("forms", {}).values()])
        if form_filter != "All":
            form_id_list = [fid for fid, f in meta["forms"].items() if f["form_name"] == form_filter]
            if form_id_list:
                responses_display = responses[responses["FormID"] == form_id_list[0]]
            else:
                responses_display = pd.DataFrame()
        else:
            responses_display = responses.copy()

        if not responses_display.empty:
            st.dataframe(responses_display, use_container_width=True)

            # Edit / Delete functionality
            for idx, row in responses_display.iterrows():
                with st.expander(f"✏️ Edit / Delete Response #{idx+1}"):
                    updated_values = {}
                    for col in meta["forms"].get(row["FormID"], {}).get("columns", []):
                        updated_values[col] = st.text_input(col, value=row[col], key=f"{col}_{idx}")

                    col1_btn, col2_btn = st.columns(2)
                    with col1_btn:
                        if st.button(f"💾 Update #{idx}", key=f"update_{idx}"):
                            for k, v in updated_values.items():
                                responses.loc[responses.index[idx], k] = v
                            save_responses(responses)
                            st.success(f"Response #{idx+1} updated!")

                    with col2_btn:
                        if st.button(f"🗑 Delete #{idx}", key=f"delete_{idx}"):
                            responses.drop(responses.index[idx], inplace=True)
                            save_responses(responses)
                            st.success(f"Response #{idx+1} deleted!")

            # Download button
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                responses_display.to_excel(writer, index=False, sheet_name="Responses")
            st.download_button(
                label="⬇️ Download Responses (Excel)",
                data=buffer.getvalue(),
                file_name="form_responses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
