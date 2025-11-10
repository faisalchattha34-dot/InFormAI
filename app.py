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
st.set_page_config(page_title="📄 Excel → Web Form", layout="centered")
st.title("📄 Auto Form Creator from Excel")

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


# ----------------------------
# URL Params
# ----------------------------
params = st.experimental_get_query_params()
mode = params.get("mode", ["admin"])[0]
form_id = params.get("form_id", [None])[0]

meta = load_meta()

# ----------------------------
# FORM VIEW (User fills form)
# ----------------------------
if mode == "form":
    if not form_id or "forms" not in meta or form_id not in meta["forms"]:
        st.error("Invalid or missing form ID. Please contact the admin.")
    else:
        info = meta["forms"][form_id]
        st.header(f"🧾 {info['form_name']}")

        # ✅ Persistent session per browser
        if "session_id" not in st.session_state:
            st.session_state["session_id"] = str(uuid.uuid4())[:8]
        session_id = st.session_state["session_id"]
        st.caption(f"🆔 Your Session ID: {session_id}")

        dropdowns = info.get("dropdowns", {})
        columns = info["columns"]

        # ✅ Create form dynamically
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
                # ✅ Ensure master Excel exists
                if not os.path.exists(ALL_RESPONSES_PATH):
                    pd.DataFrame(columns=list(row.keys())).to_excel(ALL_RESPONSES_PATH, index=False)

                existing = pd.read_excel(ALL_RESPONSES_PATH)

                # Add any new columns dynamically
                for col in row.keys():
                    if col not in existing.columns:
                        existing[col] = None

                new_row_df = pd.DataFrame([row])
                combined = pd.concat([existing, new_row_df], ignore_index=True)
                combined.to_excel(ALL_RESPONSES_PATH, index=False)

                st.success("🎉 Response saved successfully! You can add more without refreshing.")
                st.balloons()

            except Exception as e:
                st.error(f"❌ Error saving data: {e}")

# ----------------------------
# ADMIN VIEW
# ----------------------------
else:
    st.header("🧑‍💼 Admin Panel")
    st.write("Upload an Excel file — its columns will automatically become form fields.")

    uploaded = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

    if uploaded:
        try:
            df = pd.read_excel(uploaded)
            df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]
            st.success(f"Detected columns: {len(df.columns)}")
            st.write(df.columns.tolist())

            dropdowns = detect_dropdowns(uploaded, list(df.columns))
            if dropdowns:
                st.info("Detected dropdowns:")
                st.table(pd.DataFrame([{"Field": k, "Options": ", ".join(v)} for k, v in dropdowns.items()]))

            form_name = st.text_input("Form name:", value=f"My Form {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            base_url = st.text_input("Your Streamlit app public URL (example: https://yourapp.streamlit.app)")

            if st.button("🚀 Create Form Link"):
                if not base_url:
                    st.error("Please enter your app URL to generate shareable link.")
                else:
                    form_id = str(uuid.uuid4())[:10]
                    forms = meta.get("forms", {})

                    forms[form_id] = {
                        "form_name": form_name,
                        "columns": list(df.columns),
                        "dropdowns": dropdowns,
                        "created_at": datetime.now().isoformat(),
                    }

                    meta["forms"] = forms
                    save_meta(meta)

                    link = f"{base_url.rstrip('/')}/?mode=form&form_id={form_id}"
                    st.success("✅ Form created successfully!")
                    st.info("Share this link with others to fill the form:")
                    st.code(link)

                    # ----------------------------
                    # 📧 EMAIL SEND FEATURE
                    # ----------------------------
                    st.markdown("---")
                    st.subheader("📧 Send Form Link via Email")

                    sender_email = st.text_input("Your Gmail address:")
                    password = st.text_input("Your Gmail App Password (not your login password)", type="password")
                    receiver_email = st.text_input("Recipient Email:")
                    email_message = st.text_area(
                        "Optional message:",
                        f"Hi,\n\nPlease fill out this form:\n{link}\n\nThanks!"
                    )

                    if st.button("📨 Send Email"):
                        if not sender_email or not password or not receiver_email:
                            st.error("Please fill all required fields.")
                        else:
                            try:
                                msg = MIMEMultipart()
                                msg["From"] = sender_email
                                msg["To"] = receiver_email
                                msg["Subject"] = f"Form Link: {form_name}"
                                msg.attach(MIMEText(email_message, "plain"))

                                with smtplib.SMTP("smtp.gmail.com", 587) as server:
                                    server.starttls()
                                    server.login(sender_email, password)
                                    server.send_message(msg)

                                st.success(f"✅ Email sent successfully to {receiver_email}!")
                            except Exception as e:
                                st.error(f"❌ Error sending email: {e}")

        except Exception as e:
            st.error(f"Error processing file: {e}")

    st.markdown("---")
    st.subheader("📊 Existing Forms")

    forms = meta.get("forms", {})
    if forms:
        df_forms = pd.DataFrame([
            {"Form ID": fid, "Form Name": fdata["form_name"], "Created": fdata["created_at"]}
            for fid, fdata in forms.items()
        ])
        st.dataframe(df_forms)

        st.markdown("---")
        st.subheader("📈 Responses Dashboard")

        if os.path.exists(ALL_RESPONSES_PATH):
            try:
                df_responses = pd.read_excel(ALL_RESPONSES_PATH)
                if not df_responses.empty:
                    st.success(f"✅ {len(df_responses)} total responses found")
                    st.dataframe(df_responses, use_container_width=True)

                    # 🔽 Select which form to export
                    selected_form_export = st.selectbox(
                        "Select Form to Download:",
                        ["All"] + [f["form_name"] for f in forms.values()],
                        key="download_form_select"
                    )

                    export_df = df_responses.copy()
                    if selected_form_export != "All":
                        selected_form_id = [
                            fid for fid, f in forms.items() if f["form_name"] == selected_form_export
                        ][0]
                        export_df = export_df[export_df["FormID"] == selected_form_id]

                        original_columns = forms[selected_form_id]["columns"]
                        export_df = export_df[original_columns]

                    # 🔽 Download Excel version
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        export_df.to_excel(writer, index=False, sheet_name="Responses")
                    st.download_button(
                        label="⬇️ Download Responses (Excel)",
                        data=buffer.getvalue(),
                        file_name="form_responses.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                else:
                    st.info("No responses submitted yet.")
            except Exception as e:
                st.error(f"Error reading all responses: {e}")
        else:
            st.info("No responses file found yet.")
    else:
        st.info("No forms created yet.")
