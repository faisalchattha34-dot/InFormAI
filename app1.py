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
# Custom CSS Styling
# ----------------------------
st.markdown(
    """
    <style>
        body { background-color: #f4f7fc; font-family: 'Arial', sans-serif; }
        h1, h2, h3 { color: #2c3e50; font-weight: 700; }
        .stTextInput, .stSelectbox, .stTextArea, .stDataFrame {
            background-color: #ffffff !important; border-radius: 8px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.05); padding: 12px;
        }
        .stButton>button {
            background-color: #3498db; color: white; padding: 10px 20px;
            border-radius: 8px; border: none; font-size: 16px; font-weight: 500;
            transition: all 0.3s ease;
        }
        .stButton>button:hover { background-color: #2980b9; transform: scale(1.03); }
        .stDownloadButton>button {
            background-color: #2ecc71; color: white; padding: 10px 20px;
            border-radius: 8px; border: none; font-size: 16px; font-weight: 500;
            transition: all 0.3s ease;
        }
        .stDownloadButton>button:hover { background-color: #27ae60; transform: scale(1.03); }
        .container { display: flex; justify-content: space-between; gap: 20px; flex-wrap: wrap; margin-bottom: 20px; }
        .container > div { flex: 1; min-width: 300px; }
        .stTable { border-radius: 8px; background-color: white; padding: 10px; box-shadow: 0px 4px 8px rgba(0,0,0,0.05); }
    </style>
    """,
    unsafe_allow_html=True
)

# ----------------------------
# Paths and Helpers
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

            # 🧠 Intelligent Header Detection + Merged Header Handling
            excel_data = pd.read_excel(form_file, header=None)
            header_row_index = None
            for i, row in excel_data.iterrows():
                non_empty_count = row.count()
                if non_empty_count >= len(row) / 2:
                    header_row_index = i
                    break

            if header_row_index is not None:
                df_form = pd.read_excel(form_file, header=header_row_index)
            else:
                df_form = pd.read_excel(form_file)

            # 🧹 Clean + merge adjacent blank headers (handle merged cells)
            cleaned_cols = []
            prev_name = None
            for c in df_form.columns:
                name = str(c).strip() if pd.notna(c) and str(c).strip() else None
                if not name and prev_name:
                    continue  # skip merged blank header
                if name:
                    name = name.replace("_", " ").title()
                    cleaned_cols.append(name)
                    prev_name = name

            df_form = df_form.iloc[:, :len(cleaned_cols)]
            df_form.columns = cleaned_cols

            # Continue with normal logic
            if "Email" not in df_members.columns:
                st.error("❌ Member file must contain an 'Email' column.")
            else:
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
                        st.error("Please enter your app URL.")
                    elif not sender_email or not password:
                        st.error("Please enter Gmail and App Password.")
                    else:
                        form_id_new = str(uuid.uuid4())[:10]
                        forms = meta.get("forms", {})
                        forms[form_id_new] = {
                            "form_name": form_name,
                            "columns": list(df_form.columns),
                            "dropdowns": dropdowns,
                            "created_at": datetime.now().isoformat(),
                        }
                        meta["forms"] = forms
                        save_meta(meta)
                        link = f"{base_url.rstrip('/')}/?mode=form&form_id={form_id_new}"
                        st.success(f"✅ Form created successfully!\n{link}")
                        st.info("📧 Sending form link to all members...")
                        emails = df_members["Email"].dropna().unique().tolist()
                        subject = f"Form Invitation: {form_name}"
                        message = f"Hello,\n\nPlease fill out the form below:\n{link}\n\nThank you!"
                        sent_count, send_results = send_email_to_members(
                            sender_email, password, emails, subject, message
                        )
                        st.success(f"🎉 Emails sent: {sent_count}/{len(emails)}")
                        st.subheader("📧 Email Send Status")
                        st.table(pd.DataFrame(send_results))
        except Exception as e:
            st.error(f"❌ Error processing files: {e}")

    st.markdown("---")
    st.subheader("📊 Responses Dashboard")
    responses = load_responses()
    if responses.empty:
        st.info("No responses submitted yet.")
    else:
        form_filter = st.selectbox(
            "Select Form to View Responses:",
            ["All"] + [f["form_name"] for f in meta.get("forms", {}).values()],
        )
        if form_filter != "All":
            form_id_list = [
                fid for fid, f in meta["forms"].items() if f["form_name"] == form_filter
            ]
            responses_display = (
                responses[responses["FormID"] == form_id_list[0]]
                if form_id_list
                else pd.DataFrame()
            )
        else:
            responses_display = responses.copy()
        if not responses_display.empty:
            st.dataframe(responses_display, use_container_width=True)
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
            original_form_cols = []
            if form_filter != "All" and form_id_list:
                original_form_cols = meta["forms"][form_id_list[0]]["columns"]
            elif form_filter == "All":
                all_form_cols = []
                for f in meta.get("forms", {}).values():
                    all_form_cols.extend(f["columns"])
                original_form_cols = list(dict.fromkeys(all_form_cols))
            filtered_responses = (
                responses_display[original_form_cols]
                if original_form_cols
                else responses_display.copy()
            )
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                filtered_responses.to_excel(writer, index=False, sheet_name="Responses")
            st.download_button(
                label="⬇️ Download Responses (Excel)",
                data=buffer.getvalue(),
                file_name="form_responses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
