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

# Add custom CSS for styling
st.markdown(
    """
    <style>
        /* Background and Padding */
        body {
            background-color: #f4f7fc; /* Light blue background */
            padding: 20px;
            font-family: 'Arial', sans-serif;
        }

        /* Header */
        h1, h2, h3 {
            color: #2c3e50;
            font-weight: bold;
        }

        /* Form Styling */
        .stTextInput, .stSelectbox, .stButton, .stTextArea {
            background-color: #ffffff;
            border-radius: 8px;
            box-shadow: 0px 4px 8px rgba(0, 0, 0, 0.1);
            padding: 10px 15px;
            margin-bottom: 20px;
        }

        /* Submit Button */
        .stButton>button {
            background-color: #3498db;
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            font-size: 16px;
            border: none;
            cursor: pointer;
            transition: background-color 0.3s;
        }

        .stButton>button:hover {
            background-color: #2980b9;
        }

        /* Flexbox Layout */
        .container {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 20px;
        }

        .container > div {
            flex: 1;
            min-width: 300px;
        }

        /* Table Styling */
        .stTable {
            background-color: #ffffff;
            box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1);
            border-radius: 8px;
            padding: 15px;
        }

        /* Response Table Column Headers */
        .stTable thead {
            background-color: #2980b9;
            color: white;
        }

        .stTable td, .stTable th {
            padding: 12px 15px;
            text-align: left;
        }

        /* Download Button */
        .stDownloadButton>button {
            background-color: #2ecc71;
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            font-size: 16px;
            border: none;
            cursor: pointer;
            transition: background-color 0.3s;
        }

        .stDownloadButton>button:hover {
            background-color: #27ae60;
        }
    </style>
    """, unsafe_allow_html=True
)

# ----------------------------
# Helper Functions
# ----------------------------
DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")
ALL_RESPONSES_PATH = os.path.join(DATA_DIR, "all_responses.xlsx")

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
                    options = [x.strip() for x in formula.split(",")] if "," in formula else []
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

    # Flexbox layout for the file uploaders
    st.markdown('<div class="container">', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        member_file = st.file_uploader("📋 Upload Member List (must have 'Email' column)", type=["xlsx"])
    with col2:
        form_file = st.file_uploader("📄 Upload Form Source File", type=["xlsx"])
    
    st.markdown('</div>', unsafe_allow_html=True)

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
                sender
