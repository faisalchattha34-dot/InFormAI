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
st.title("📄 Excel → Web Form + Auto Email SaaS")

# Folders
os.makedirs("forms", exist_ok=True)
os.makedirs("responses", exist_ok=True)
if not os.path.exists("meta.json"):
    with open("meta.json", "w") as f:
        json.dump({"forms": {}}, f)

# ----------------------------
# Load Meta
# ----------------------------
def load_meta():
    with open("meta.json", "r") as f:
        return json.load(f)

def save_meta(meta):
    with open("meta.json", "w") as f:
        json.dump(meta, f, indent=4)

meta = load_meta()

# ----------------------------
# Helper Functions
# ----------------------------
def send_email(receiver_email, form_link):
    sender_email = "your_email@example.com"
    sender_pass = "your_email_password"
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = "Please fill the form"
    body = f"Please fill this form: {form_link}"
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_pass)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Email failed: {e}")
        return False

def load_responses():
    all_files = os.listdir("responses")
    all_data = []
    for file in all_files:
        with open(f"responses/{file}", "r") as f:
            all_data.append(json.load(f))
    if all_data:
        return pd.DataFrame(all_data)
    else:
        return pd.DataFrame(columns=["form_id", "response_id", "timestamp", "data"])

# ----------------------------
# 1️⃣ Create New Form
# ----------------------------
st.sidebar.header("📋 Create / Manage Forms")
form_name = st.sidebar.text_input("Form Name")
uploaded_file = st.sidebar.file_uploader("Upload Excel (Columns will become fields)", type=["xlsx"])

if st.sidebar.button("Create Form") and form_name and uploaded_file:
    df = pd.read_excel(uploaded_file)
    columns = df.columns.tolist()
    form_id = str(uuid.uuid4())
    
    meta["forms"][form_id] = {
        "form_name": form_name,
        "columns": columns,
        "created_at": str(datetime.now())
    }
    save_meta(meta)
    
    with open(f"forms/{form_id}.json", "w") as f:
        json.dump({"fields": columns}, f, indent=4)
    
    st.success(f"Form '{form_name}' created with fields: {columns}")

# ----------------------------
# 2️⃣ Share Form (Email)
# ----------------------------
st.sidebar.subheader("📧 Share Form via Email")
if meta.get("forms"):
    selected_form = st.sidebar.selectbox(
        "Select Form to Share", 
        [(f["form_name"], fid) for fid, f in meta["forms"].items()]
    )
    if selected_form:
        form_label, form_id = selected_form
        emails_text = st.sidebar.text_area("Enter emails (comma separated)")
        if st.sidebar.button("Send Form"):
            emails = [e.strip() for e in emails_text.split(",") if e.strip()]
            link = f"http://localhost:8501/?form_id={form_id}"  # Update for deployed URL
            for email in emails:
                send_email(email, link)
            st.sidebar.success(f"Form sent to {len(emails)} emails.")

# ----------------------------
# 3️⃣ Fill Form
# ----------------------------
query_params = st.experimental_get_query_params()
if "form_id" in query_params:
    fid = query_params["form_id"][0]
    form = meta["forms"].get(fid)
    if form:
        st.subheader(f"📝 {form['form_name']}")
        fields = form["columns"]
        user_data = {}
        for f in fields:
            user_data[f] = st.text_input(f)
        if st.button("Submit"):
            response_id = str(uuid.uuid4())
            response_data = {
                "form_id": fid,
                "response_id": response_id,
                "timestamp": str(datetime.now()),
                "data": user_data
            }
            with open(f"responses/{response_id}.json", "w") as f:
                json.dump(response_data, f, indent=4)
            st.success("Response submitted successfully!")
            st.experimental_rerun()

# ----------------------------
# 4️⃣ Responses Dashboard
# ----------------------------
st.markdown("---")
st.subheader("📊 Responses Dashboard")
responses = load_responses()

# Initialize session state flags
if "update_rerun" not in st.session_state:
    st.session_state.update_rerun = False

if responses.empty:
    st.info("No responses submitted yet.")
else:
    form_filter = st.selectbox(
        "Select Form to View Responses:", 
        ["All"] + [f["form_name"] for f in meta.get("forms", {}).values()]
    )
    if form_filter != "All":
        fids = [fid for fid, f in meta["forms"].items() if f["form_name"] == form_filter]
        responses = responses[responses["form_id"].isin(fids)]
    
    # Safe edit/delete loop
    for idx, row in responses.iterrows():
        st.markdown(f"**Response ID:** {row['response_id']}  |  Submitted: {row['timestamp']}")
        data_dict = row["data"].copy()  # copy to avoid overwriting original
        
        cols = st.columns([3,1,1])
        with cols[0]:
            # Editable fields
            for key in data_dict:
                data_dict[key] = st.text_input(
                    f"{key}", 
                    value=data_dict[key], 
                    key=f"{row['response_id']}_{key}"
                )

        with cols[1]:
            save_btn = st.button("Save", key=f"save_{row['response_id']}")
            if save_btn:
                response_data = {
                    "form_id": row["form_id"],
                    "response_id": row["response_id"],
                    "timestamp": row["timestamp"],
                    "data": data_dict
                }
                with open(f"responses/{row['response_id']}.json", "w") as f:
                    json.dump(response_data, f, indent=4)
                st.session_state.update_rerun = True

        with cols[2]:
            delete_btn = st.button("Delete", key=f"delete_{row['response_id']}")
            if delete_btn:
                os.remove(f"responses/{row['response_id']}.json")
                st.session_state.update_rerun = True

# Rerun outside loop once
if st.session_state.update_rerun:
    st.session_state.update_rerun = False
    st.experimental_rerun()
