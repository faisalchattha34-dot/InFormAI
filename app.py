# app.py
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
# Paths & Helpers
# ----------------------------
DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")
ALL_RESPONSES_PATH = os.path.join(DATA_DIR, "all_responses.xlsx")

def load_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"forms": {}}

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

def send_email(sender_email, sender_password, to_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

def load_responses():
    if os.path.exists(ALL_RESPONSES_PATH):
        return pd.read_excel(ALL_RESPONSES_PATH)
    return pd.DataFrame()

def save_responses(df):
    df.to_excel(ALL_RESPONSES_PATH, index=False)

meta = load_meta()

# ----------------------------
# Sidebar / Admin Email settings for notifications
# ----------------------------
st.sidebar.header("🔧 Admin / Email settings")
st.sidebar.write("These SMTP credentials are used to notify respondents when you edit a response.")
sender_email_setting = st.sidebar.text_input("Notify sender email (Gmail)", key="notify_sender")
sender_pass_setting = st.sidebar.text_input("Notify sender app password", type="password", key="notify_pass")
enable_notify_on_edit = st.sidebar.checkbox("Notify respondent by email when edited?", value=True)

# ----------------------------
# URL Params
# ----------------------------
params = st.experimental_get_query_params()
mode = params.get("mode", ["admin"])[0]
form_id = params.get("form_id", [None])[0]

# ----------------------------
# FORM VIEW (for end-users)
# ----------------------------
if mode == "form":
    if not form_id or "forms" not in meta or form_id not in meta["forms"]:
        st.warning("Invalid or missing form ID. Please open a form link created by admin.")
        if "forms" in meta and meta["forms"]:
            st.markdown("Available forms:")
            for fid, f in meta["forms"].items():
                st.markdown(f"- {f['form_name']} (id: {fid})")
    else:
        info = meta["forms"][form_id]
        st.header(f"🧾 {info['form_name']}")
        # session id to avoid overwriting
        if "session_id" not in st.session_state:
            st.session_state["session_id"] = str(uuid.uuid4())[:8]
        session_id = st.session_state["session_id"]

        dropdowns = info.get("dropdowns", {})
        columns = info["columns"]
        # show form
        with st.form("user_form", clear_on_submit=False):
            values = {}
            for col in columns:
                if col in dropdowns and isinstance(dropdowns[col], list) and dropdowns[col]:
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
    st.write("Upload two Excel files — Member List & Form Source. (Members => who receive link; Form => columns for the form.)")

    col1, col2 = st.columns(2)
    with col1:
        member_file = st.file_uploader("📋 Upload Member List (must have 'Email' column)", type=["xlsx"])
    with col2:
        form_file = st.file_uploader("📄 Upload Form Source File", type=["xlsx"])

    if member_file and form_file:
        try:
            # read members
            df_members = pd.read_excel(member_file)
            # attempt to find header row for form source
            excel_data = pd.read_excel(form_file, header=None)
            header_row_index = None
            for i, row in excel_data.iterrows():
                if row.count() >= len(row) / 2:
                    header_row_index = i
                    break
            df_form = pd.read_excel(form_file, header=header_row_index if header_row_index is not None else 0)

            # Clean header names (like before)
            cleaned_cols = []
            seen = set()
            prev_name = None
            for c in df_form.columns:
                name = str(c).strip() if pd.notna(c) and str(c).strip() else prev_name
                if name:
                    name = name.replace("_", " ").title()
                    if name in seen:
                        i = 2
                        while f"{name}_{i}" in seen:
                            i += 1
                        name = f"{name}_{i}"
                    seen.add(name)
                    cleaned_cols.append(name)
                    prev_name = name
            df_form.columns = cleaned_cols

            # Show editable preview of form columns
            st.subheader("👀 Edit Form Data (Live Preview)")
            if "original_columns" not in st.session_state:
                st.session_state.original_columns = list(df_form.columns)
            if "current_form_df" not in st.session_state:
                st.session_state.current_form_df = df_form.copy()

            edited_df = st.data_editor(
                st.session_state.current_form_df,
                use_container_width=True,
                num_rows="dynamic",
                key="form_editor",
            )
            st.session_state.current_form_df = edited_df.copy()

            st.write("### ✏️ Column Management")
            col_action = st.radio("Select Action", ["None", "Rename Column", "Delete Column", "Add Column", "Restore Deleted Column"], horizontal=True)

            if col_action == "Rename Column":
                col_to_rename = st.selectbox("Select column to rename", st.session_state.current_form_df.columns)
                new_name = st.text_input("Enter new column name:")
                if st.button("✅ Rename Now"):
                    st.session_state.current_form_df.rename(columns={col_to_rename: new_name}, inplace=True)
                    st.success(f"Column renamed from '{col_to_rename}' → '{new_name}'")

            elif col_action == "Delete Column":
                col_to_delete = st.selectbox("Select column to delete", st.session_state.current_form_df.columns)
                if st.button("🗑️ Delete Column"):
                    st.session_state.current_form_df.drop(columns=[col_to_delete], inplace=True)
                    st.success(f"Column '{col_to_delete}' deleted.")

            elif col_action == "Add Column":
                new_col_name = st.text_input("Enter new column name:")
                if st.button("➕ Add Column"):
                    if new_col_name in st.session_state.current_form_df.columns:
                        st.warning("Column already exists.")
                    else:
                        st.session_state.current_form_df[new_col_name] = ""
                        st.success(f"Column '{new_col_name}' added.")

            elif col_action == "Restore Deleted Column":
                deleted_cols = [c for c in st.session_state.original_columns if c not in st.session_state.current_form_df.columns]
                if deleted_cols:
                    col_to_restore = st.selectbox("Select deleted column to restore", deleted_cols)
                    if st.button("♻️ Restore Column"):
                        idx = st.session_state.original_columns.index(col_to_restore)
                        df = st.session_state.current_form_df
                        df.insert(loc=idx, column=col_to_restore, value="")
                        st.session_state.current_form_df = df
                        st.success(f"Column '{col_to_restore}' restored successfully.")
                else:
                    st.info("No deleted columns found to restore.")

            save_changes = st.button("💾 Save Changes to Original Excel File")
            if save_changes:
                try:
                    with BytesIO() as buffer:
                        st.session_state.current_form_df.to_excel(buffer, index=False)
                        buffer.seek(0)
                        # overwrite uploaded file locally (note: this just writes a file with same name on server)
                        with open(form_file.name, "wb") as f:
                            f.write(buffer.read())
                    st.success("✅ All changes saved back to the uploaded Excel file successfully!")
                except Exception as e:
                    st.error(f"❌ Failed to save: {e}")

            # Continue workflow (detect dropdowns, prepare to create form)
            if "Email" not in df_members.columns:
                st.error("❌ Member file must contain an 'Email' column.")
            else:
                dropdowns = detect_dropdowns(form_file, list(st.session_state.current_form_df.columns))
                st.success(f"✅ Form fields detected: {len(st.session_state.current_form_df.columns)}")
                st.write("Columns:", st.session_state.current_form_df.columns.tolist())
                if dropdowns:
                    st.info("Detected dropdowns:")
                    st.table(pd.DataFrame([{"Field": k, "Options": ", ".join(v)} for k, v in dropdowns.items()]))

                form_name = st.text_input("Form Name:", value=f"My Form {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                base_url = st.text_input("Your Streamlit App Public URL (example: https://yourapp.streamlit.app)")
                sender_email = st.text_input("Your Gmail Address (for sending links):")
                password = st.text_input("Your Gmail App Password (for sending links):", type="password")

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
                            "columns": list(st.session_state.current_form_df.columns),
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

                        sent_count, send_results = 0, []
                        for e in emails:
                            ok, err = send_email(sender_email, password, e, subject, message)
                            if ok:
                                sent_count += 1
                                send_results.append({"Email": e, "Status": "✅ Sent"})
                            else:
                                send_results.append({"Email": e, "Status": f"❌ Failed ({err})"})
                        st.success(f"🎉 Emails attempted: {sent_count}/{len(emails)}")
                        st.subheader("📧 Email Send Status")
                        st.table(pd.DataFrame(send_results))

        except Exception as e:
            st.error(f"❌ Error processing files: {e}")

    # ----------------------------
    # Responses Dashboard
    # ----------------------------
    st.markdown("---")
    st.subheader("📊 Responses Dashboard (Search / Pagination / Edit / Delete)")

    # Load responses
    responses = load_responses()
    if responses.empty:
        st.info("No responses submitted yet.")
    else:
        # Search
        search_query = st.text_input("🔎 Search (any column):")
        df_filtered = responses.copy()
        if search_query:
            df_filtered = df_filtered[df_filtered.apply(lambda r: r.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)]

        # Form filter dropdown
        form_filter = st.selectbox("Filter by Form:", ["All"] + [f["form_name"] for f in meta.get("forms", {}).values()])
        if form_filter != "All":
            form_id_list = [fid for fid, f in meta["forms"].items() if f["form_name"] == form_filter]
            if form_id_list:
                df_filtered = df_filtered[df_filtered["FormID"] == form_id_list[0]]

        # Pagination
        PAGE_SIZE = st.number_input("Rows per page:", min_value=5, max_value=100, value=10, step=5)
        total = len(df_filtered)
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        if "page" not in st.session_state:
            st.session_state.page = 1
        page_col1, page_col2, page_col3 = st.columns([1, 6, 1])
        if page_col1.button("⬅ Prev") and st.session_state.page > 1:
            st.session_state.page -= 1
        page_col2.write(f"Page {st.session_state.page} / {total_pages}  —  Total rows: {total}")
        if page_col3.button("Next ➡") and st.session_state.page < total_pages:
            st.session_state.page += 1

        start = (st.session_state.page - 1) * PAGE_SIZE
        end = start + PAGE_SIZE
        page_df = df_filtered.iloc[start:end]

        # Show page table (hide metadata columns optionally)
        hide_meta = st.checkbox("Hide metadata columns (FormID, FormName, UserSession, SubmittedAt)", value=True)
        display_df = page_df.copy()
        if hide_meta:
            for c in ["FormID", "FormName", "UserSession", "SubmittedAt"]:
                if c in display_df.columns:
                    display_df = display_df.drop(columns=[c])

        st.dataframe(display_df, use_container_width=True)

        # Edit specific response (show form UI like original)
        st.markdown("### ✏️ Edit a response (opens below)")
        if not page_df.empty:
            edit_row_number = st.number_input("Select row index in current page to edit (0 = first row shown)", min_value=0, max_value=max(0, len(page_df) - 1), value=0, step=1)
            if st.button("✍️ Open edit form for selected row"):
                # map page index to real index
                real_index = page_df.index[edit_row_number]
                row_to_edit = responses.loc[real_index]
                selected_form_id = row_to_edit["FormID"]
                form_info = meta["forms"].get(selected_form_id, None)
                if not form_info:
                    st.error("Form definition not found for this response (maybe form was deleted).")
                else:
                    st.subheader(f"Editing response for: {form_info['form_name']}")
                    dropdowns = form_info.get("dropdowns", {})
                    columns = form_info.get("columns", [])

                    with st.form("edit_response_form"):
                        edited_values = {}
                        for col in columns:
                            current_val = row_to_edit.get(col, "")
                            if col in dropdowns and isinstance(dropdowns[col], list) and dropdowns[col]:
                                idx = 0
                                try:
                                    idx = dropdowns[col].index(current_val) if current_val in dropdowns[col] else 0
                                except Exception:
                                    idx = 0
                                edited_values[col] = st.selectbox(col, dropdowns[col], index=idx)
                            else:
                                edited_values[col] = st.text_input(col, value=str(current_val))

                        save_edit = st.form_submit_button("💾 Save Edited Response")

                    if save_edit:
                        # Write edits back to responses (preserve metadata columns)
                        for col, val in edited_values.items():
                            responses.at[real_index, col] = val

                        # Update modified timestamp column or create one
                        responses.at[real_index, "LastEditedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        save_responses(responses)
                        st.success("✅ Edited response saved.")

                        # Notify respondent by email if enabled and we have their email
                        if enable_notify_on_edit and sender_email_setting and sender_pass_setting:
                            respondent_email = responses.at[real_index, "Email"] if "Email" in responses.columns else None
                            if respondent_email and pd.notna(respondent_email):
                                subj = f"Your response to '{form_info['form_name']}' was updated"
                                body_lines = [
                                    f"Hello,",
                                    "",
                                    f"Your submitted response to the form '{form_info['form_name']}' was updated by the admin.",
                                    "If you did not request this, please contact the administrator.",
                                    "",
                                    "Best regards,"
                                ]
                                body = "\n".join(body_lines)
                                ok, err = send_email(sender_email_setting, sender_pass_setting, respondent_email, subj, body)
                                if ok:
                                    st.info(f"Notification email sent to {respondent_email}")
                                else:
                                    st.warning(f"Failed to send notification email to {respondent_email}: {err}")
                            else:
                                st.info("No respondent email found in this response to send notification to.")

                        st.experimental_rerun()

        # Delete response
        st.markdown("### 🗑️ Delete a response")
        if not page_df.empty:
            del_row_number = st.number_input("Select row index in current page to delete (0 = first row shown)", min_value=0, max_value=max(0, len(page_df) - 1), value=0, step=1, key="del_idx")
            if st.button("❌ Delete selected response"):
                real_index = page_df.index[del_row_number]
                responses = responses.drop(real_index).reset_index(drop=True)
                save_responses(responses)
                st.success("🗑️ Response deleted.")
                st.experimental_rerun()

        # Download displayed (filtered) responses
        st.markdown("---")
        st.write("Download visible responses:")
        to_download = BytesIO()
        # write the currently filtered full df (not only page) for convenience
        df_filtered.to_excel(to_download, index=False)
        to_download.seek(0)
        st.download_button(
            "📥 Download Filtered Responses",
            data=to_download,
            file_name="filtered_responses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

