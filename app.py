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

# ----------------------------
# CSS (UI Improve)
# ----------------------------
st.markdown("""
<style>
:root { color-scheme: light dark; }
.stButton>button { background-color:#3498db;color:white;padding:10px 18px;border:none;border-radius:6px; }
.stButton>button:hover { background-color:#2980b9; }
</style>
""", unsafe_allow_html=True)

# ----------------------------
# Paths
# ----------------------------
DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")
ALL_RESPONSES_PATH = os.path.join(DATA_DIR, "all_responses.xlsx")

# ----------------------------
# Meta Functions
# ----------------------------
def load_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_meta(meta):
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

# ----------------------------
# FIXED DROPDOWN FUNCTION
# ----------------------------
def detect_dropdowns(excel_file, df_columns):
    excel_file.seek(0)
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active
    dropdowns = {}

    if not ws.data_validations:
        return dropdowns

    for dv in ws.data_validations.dataValidation:
        try:
            if dv.type != "list" or not dv.formula1:
                continue

            formula = str(dv.formula1).strip()

            # Inline Values e.g "Male,Female"
            if "," in formula:
                options = [x.strip().strip('"') for x in formula.strip('"').split(",")]

            # Range e.g =Sheet1!$B$3:$B$10
            else:
                try:
                    ref = formula.replace("=", "").replace("$", "")
                    sheet_name, cells = ref.split("!")
                    col_range = cells.split(":")

                    ws2 = wb[sheet_name]
                    start = col_range[0]
                    end = col_range[-1]
                    options = []

                    for row in ws2[start:end]:
                        for cell in row:
                            if cell.value not in [None, ""]:
                                options.append(str(cell.value))
                except:
                    continue

            # Assign dropdown to field column
            for cell_range in dv.cells:
                col_index = cell_range.min_col - 1
                if 0 <= col_index < len(df_columns):
                    dropdowns[df_columns[col_index]] = sorted(list(set(options)))

        except:
            continue

    return dropdowns

# ----------------------------
# Send Email
# ----------------------------
def send_email_to_members(sender_email,password,members,subject,message):
    sent=0
    out=[]
    for email in members:
        try:
            msg=MIMEMultipart()
            msg["From"]=sender_email
            msg["To"]=email
            msg["Subject"]=subject
            msg.attach(MIMEText(message,"plain"))
            with smtplib.SMTP("smtp.gmail.com",587) as server:
                server.starttls()
                server.login(sender_email,password)
                server.send_message(msg)
            sent+=1
            out.append({"Email":email,"Status":"Sent"})
        except:
            out.append({"Email":email,"Status":"Failed"})
    return sent,out

# ----------------------------
# Responses Save/Load
# ----------------------------
def load_responses():
    if os.path.exists(ALL_RESPONSES_PATH):
        return pd.read_excel(ALL_RESPONSES_PATH)
    return pd.DataFrame()

def save_responses(df):
    df.to_excel(ALL_RESPONSES_PATH,index=False)

# ----------------------------
# URL Modes
# ----------------------------
params = st.experimental_get_query_params()
mode = params.get("mode",["admin"])[0]
form_id = params.get("form_id",[None])[0]
meta = load_meta()

# ===================================================================
#                        FORM PAGE
# ===================================================================
if mode=="form":
    if not form_id or "forms" not in meta or form_id not in meta["forms"]:
        st.warning("Invalid Form Link.")
    else:
        info = meta["forms"][form_id]
        st.header(f"📝 {info['form_name']}")

        if "session_id" not in st.session_state:
            st.session_state["session_id"]=str(uuid.uuid4())[:8]
        sid=st.session_state["session_id"]

        dropdowns = info.get("dropdowns",{})
        columns=info["columns"]

        with st.form("user_form"):
            values={}
            for col in columns:
                if col in dropdowns:
                    values[col]=st.selectbox(col,dropdowns[col])
                else:
                    values[col]=st.text_input(col)
            submitted=st.form_submit_button("Submit")

        if submitted:
            row={
                "FormID":form_id,
                "SubmittedAt":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            row.update(values)
            responses=load_responses()
            for c in row.keys():
                if c not in responses.columns:
                    responses[c]=None
            responses=pd.concat([responses,pd.DataFrame([row])],ignore_index=True)
            save_responses(responses)
            st.success("✔ Response Saved")
            st.balloons()

# ===================================================================
#                        ADMIN PAGE
# ===================================================================
else:
    st.header("🧑‍💼 Admin Panel")

    col1,col2 = st.columns(2)
    with col1:
        member_file=st.file_uploader("Member List (Email required)", type=["xlsx"])
    with col2:
        form_file=st.file_uploader("Form Source", type=["xlsx"])

    if member_file and form_file:
        df_members=pd.read_excel(member_file)
        df_form=pd.read_excel(form_file)

        st.subheader("Edit Form Columns")
        edited=st.data_editor(df_form, num_rows="dynamic", key="edit")
        dropdowns=detect_dropdowns(form_file, list(edited.columns))

        st.write("Dropdowns Found:")
        st.write(dropdowns)

        form_name=st.text_input("Form Name")
        base_url=st.text_input("Streamlit Public URL")
        sender_email=st.text_input("Email")
        password=st.text_input("App Password", type="password")

        if st.button("Create Form & Send Emails"):
            new_id=str(uuid.uuid4())[:10]
            meta.setdefault("forms",{})[new_id]={
                "form_name":form_name,
                "columns":list(edited.columns),
                "dropdowns":dropdowns,
            }
            save_meta(meta)
            link=f"{base_url}/?mode=form&form_id={new_id}"

            emails=df_members["Email"].dropna().unique().tolist()
            subject="Please fill this form"
            msg=f"Open form: {link}"

            sent,out=send_email_to_members(sender_email,password,emails,subject,msg)
            st.write(f"Sent {sent}/{len(emails)}")
            st.table(pd.DataFrame(out))

    # -----------------------
    st.subheader("Responses")
    resp=load_responses()
    if not resp.empty:
        st.dataframe(resp,use_container_width=True)
