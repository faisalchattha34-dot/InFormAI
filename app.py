# app.py
import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime
from io import BytesIO
from pathlib import Path
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import urllib.parse

# ----------------------------
# Configuration / Storage
# ----------------------------
st.set_page_config(page_title="Excel → Dynamic Form System", layout="wide")
BASE = Path("data_store")
TEMPLATES_DIR = BASE / "templates"
MEMBERS_DIR = BASE / "members"
RESPONSES_DIR = BASE / "responses"
SENT_EMAILS_LOG = BASE / "sent_emails.json"
for d in (TEMPLATES_DIR, MEMBERS_DIR, RESPONSES_DIR):
    d.mkdir(parents=True, exist_ok=True)

# ----------------------------
# Helpers
# ----------------------------
def save_json(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, ensure_ascii=False)

def load_json(path: Path):
    if not path.exists():
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def send_email(smtp_server, smtp_port, sender_email, sender_pass, receiver_email, subject, body):
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=20)
        server.starttls()
        server.login(sender_email, sender_pass)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        return True, None
    except Exception as e:
        return False, str(e)

def detect_template_from_file(uploaded_file):
    """
    Accepts an uploaded file (BytesIO or UploadedFile) and returns:
    - columns: list of column names
    - dropdowns: dict field -> [options] (detected)
    Detection strategies:
      1) If there's a sheet named 'Dropdowns' with two columns: Field, Options (comma-separated)
      2) If any header cell contains 'Field: a,b,c' (inline) -> parse
      3) Otherwise treat headers as simple text fields
    """
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception:
        # try csv
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file)
        return list(df.columns), {}
    dropdowns = {}
    columns = []

    # strategy 1: look for "Dropdowns" sheet
    if "Dropdowns" in xls.sheet_names:
        df_dd = pd.read_excel(uploaded_file, sheet_name="Dropdowns", engine="openpyxl")
        # Expect columns: Field, Options
        for _, row in df_dd.iterrows():
            if pd.isna(row.iloc[0]): 
                continue
            field = str(row.iloc[0]).strip()
            opts = str(row.iloc[1]) if len(row) > 1 else ""
            opt_list = [o.strip() for o in str(opts).split(",") if o.strip()]
            if opt_list:
                dropdowns[field] = opt_list

    # primary sheet for columns: pick first sheet that's not 'Dropdowns'
    main_sheet_name = next((s for s in xls.sheet_names if s != "Dropdowns"), xls.sheet_names[0])
    df_main = pd.read_excel(uploaded_file, sheet_name=main_sheet_name, engine="openpyxl")

    # strategy 2: headers may contain inline options like "Gender: Male,Female"
    for raw_col in df_main.columns:
        raw = str(raw_col)
        if ":" in raw and "," in raw:
            # inline dropdown
            name, opts = raw.split(":", 1)
            name = name.strip()
            opt_list = [o.strip() for o in opts.split(",") if o.strip()]
            dropdowns[name] = opt_list
            columns.append(name)
        else:
            columns.append(raw.strip())

    # strategy 3: if some columns have few unique values -> treat as dropdown heuristically
    # only if dropdowns not already provided for that column
    for col in columns:
        if col in dropdowns:
            continue
        if col in df_main.columns:
            vals = df_main[col].dropna().astype(str).unique()
            if 1 < len(vals) <= 20:
                # treat as dropdown candidate but only if values are short and repeated
                dropdowns[col] = list(vals)

    return columns, dropdowns

def create_member_links(form_id, members_df, form_base_url, expire_days=365):
    """
    For each member, create a unique token and link.
    Return list of dicts: {name,email,token,link}
    """
    out = []
    for _, row in members_df.iterrows():
        name = str(row.get("Name") or row.get("name") or "").strip()
        email = str(row.get("Email") or row.get("email") or "").strip()
        if not email:
            continue
        token = uuid.uuid4().hex
        params = {"form_id": form_id, "email": email, "token": token}
        link = f"{form_base_url}?{urllib.parse.urlencode(params)}"
        out.append({"name": name, "email": email, "token": token, "link": link})
    return out

def get_request_params():
    query = st.experimental_get_query_params()
    form_id = query.get("form_id", [None])[0]
    email = query.get("email", [None])[0]
    token = query.get("token", [None])[0]
    return form_id, email, token

# ----------------------------
# Sidebar - Email credentials & Mode
# ----------------------------
st.sidebar.header("🔧 Settings")
st.sidebar.info("Provide sender email (Gmail recommended) and app password to enable sending links & notifications.")
smtp_server = st.sidebar.text_input("SMTP Server", value="smtp.gmail.com")
smtp_port = st.sidebar.number_input("SMTP Port", value=587)
sender_email = st.sidebar.text_input("Sender Email (from)")
sender_pass = st.sidebar.text_input("Sender App Password", type="password")
base_url = st.sidebar.text_input("Base Form URL (app URL where members will open form)", value="http://localhost:8501")
st.sidebar.markdown("---")
st.sidebar.header("System Info")
st.sidebar.write(f"Storage: `{BASE.resolve()}`")

# ----------------------------
# Main UI Tabs
# ----------------------------
tabs = st.tabs(["Create Form & Send Links", "Open Form / Fill", "Dashboard / Responses"])
tab_create, tab_fill, tab_dash = tabs

# ----------------------------
# Tab 1: Create Template, upload members, configure and send links
# ----------------------------
with tab_create:
    st.header("1) Upload Template & Members → Build Form")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Upload Members File (xlsx or csv)")
        members_file = st.file_uploader("Members File (must contain Name & Email columns)", type=["xlsx", "csv"], key="members_upload")
        members_df = None
        if members_file:
            try:
                if members_file.type.endswith("csv") or members_file.name.lower().endswith(".csv"):
                    members_df = pd.read_csv(members_file)
                else:
                    members_df = pd.read_excel(members_file, engine="openpyxl")
                st.success("Members loaded")
                st.dataframe(members_df.head(50))
            except Exception as e:
                st.error(f"Error reading members file: {e}")

    with col2:
        st.subheader("Upload Template File (form columns + dropdowns)")
        template_file = st.file_uploader("Template File (xlsx)", type=["xlsx"], key="template_upload")
        detected_columns, detected_dropdowns = None, {}
        if template_file:
            try:
                detected_columns, detected_dropdowns = detect_template_from_file(template_file)
                st.success("Template detected")
                st.markdown("**Detected columns:**")
                st.write(detected_columns)
                if detected_dropdowns:
                    st.markdown("**Detected dropdowns:**")
                    st.json(detected_dropdowns)
            except Exception as e:
                st.error(f"Error reading template file: {e}")

    st.markdown("---")
    st.subheader("Form Configuration")
    form_title = st.text_input("Form Title", value="Dynamic Form")
    description = st.text_area("Form description (shown to respondents)", value="Please fill this form.")
    # Build an editable columns config
    config_cols = []
    if detected_columns:
        # initialize session state config if not present
        if "form_config_tmp" not in st.session_state:
            cfg = []
            for c in detected_columns:
                cfg.append({
                    "name": c,
                    "type": ("dropdown" if c in detected_dropdowns else "text"),
                    "options": detected_dropdowns.get(c, []),
                    "removed": False
                })
            st.session_state.form_config_tmp = cfg

    if "form_config_tmp" not in st.session_state:
        st.session_state.form_config_tmp = []

    # UI to edit columns
    df_cfg = pd.DataFrame(st.session_state.form_config_tmp)
    st.write("Edit columns (add / remove / set type / options). Changes saved locally in session.")
    edit_col = st.selectbox("Select column to edit (or choose Add New)", ["__ADD_NEW__"] + [c["name"] for c in st.session_state.form_config_tmp])
    if edit_col == "__ADD_NEW__":
        new_name = st.text_input("New column name")
        new_type = st.selectbox("Type", ["text", "dropdown"], key="new_type")
        new_opts = st.text_area("Options (comma separated, for dropdown only)", key="new_opts")
        if st.button("Add Column"):
            if new_name.strip():
                item = {"name": new_name.strip(), "type": new_type, "options": [o.strip() for o in new_opts.split(",") if o.strip()] if new_type=="dropdown" else [], "removed": False}
                st.session_state.form_config_tmp.append(item)
                st.success("Added")
            else:
                st.error("Provide a name")
            st.experimental_rerun()
    else:
        # edit existing
        idx = next((i for i,c in enumerate(st.session_state.form_config_tmp) if c["name"]==edit_col), None)
        if idx is not None:
            item = st.session_state.form_config_tmp[idx]
            with st.form(f"edit_form_{idx}"):
                item_name = st.text_input("Column name", value=item["name"])
                item_type = st.selectbox("Type", ["text","dropdown"], index=0 if item["type"]=="text" else 1)
                opts_str = st.text_area("Options (comma separated)", value=",".join(item.get("options",[])))
                remove_flag = st.checkbox("Mark as Removed (hide from form)", value=item.get("removed",False))
                submitted = st.form_submit_button("Save Column")
                if submitted:
                    item["name"] = item_name.strip()
                    item["type"] = item_type
                    item["options"] = [o.strip() for o in opts_str.split(",") if o.strip()] if item_type=="dropdown" else []
                    item["removed"] = remove_flag
                    st.session_state.form_config_tmp[idx] = item
                    st.success("Saved")
                    st.experimental_rerun()

    st.markdown("**Current columns config**")
    st.table(pd.DataFrame(st.session_state.form_config_tmp))

    st.markdown("---")
    st.subheader("Finalize Form & Send Links")
    if st.button("Create Form (generate links for members)"):
        if not members_df is None and template_file:
            form_id = uuid.uuid4().hex
            # assemble config
            final_cols = [c for c in st.session_state.form_config_tmp if not c.get("removed", False)]
            form_obj = {
                "id": form_id,
                "title": form_title,
                "description": description,
                "columns": final_cols,
                "created_at": datetime.utcnow().isoformat()
            }
            save_json(TEMPLATES_DIR / f"{form_id}.json", form_obj)
            st.success(f"Form created with id `{form_id}`. Now generating links and sending emails...")

            # create member links
            members_df_for_links = members_df.copy()
            links_info = create_member_links(form_id, members_df_for_links, base_url)
            # Save members with tokens
            members_store_path = MEMBERS_DIR / f"{form_id}_members.json"
            save_json(members_store_path, links_info)

            # attempt to send emails
            sent_log = load_json(SENT_EMAILS_LOG) or []
            send_results = []
            for m in links_info:
                body = f"Assalamualaikum {m['name'] or ''},\n\nYou have been invited to fill the form: {form_title}\n\nPlease open the link below to fill the form:\n\n{m['link']}\n\nIf you need to edit your response later, use the same link.\n\nRegards."
                if not sender_email or not sender_pass:
                    send_results.append({"email": m["email"], "status": "skipped", "reason": "no sender creds"})
                    continue
                ok, err = send_email(smtp_server, smtp_port, sender_email, sender_pass, m["email"], f"Fill Form: {form_title}", body)
                send_results.append({"email": m["email"], "status": "sent" if ok else "failed", "error": err})
                sent_log.append({"form_id": form_id, "to": m["email"], "status": "sent" if ok else "failed", "error": err, "time": datetime.utcnow().isoformat()})
            save_json(SENT_EMAILS_LOG, sent_log)
            st.write(pd.DataFrame(send_results))
            st.success("Done. Links saved and emails attempted. You can view members & tokens in 'Dashboard / Responses' tab.")
        else:
            st.error("Upload both members file and template file first.")

# ----------------------------
# Tab 2: Open Form / Fill (used by members via link)
# ----------------------------
with tab_fill:
    st.header("2) Open Form (fill or edit)")
    st.info("Open this tab via the personalized link sent in email. Link includes `form_id`, `email` and `token` query parameters.")

    # Get params from query
    form_id_q, email_q, token_q = get_request_params()
    st.write("Detected query params (for demo):", {"form_id": form_id_q, "email": email_q, "token": token_q})

    if not form_id_q:
        st.warning("No `form_id` in URL. To test, click a generated link from the Create tab or manually add `?form_id=...&email=...&token=...` to the URL.")
    else:
        # load template
        template_path = TEMPLATES_DIR / f"{form_id_q}.json"
        template = load_json(template_path)
        if not template:
            st.error("Form not found.")
        else:
            st.subheader(template.get("title", "Form"))
            st.write(template.get("description", ""))
            # load members to validate token (optional)
            members_path = MEMBERS_DIR / f"{form_id_q}_members.json"
            members = load_json(members_path) or []
            member_entry = next((m for m in members if m["email"]==email_q and m["token"]==token_q), None) if email_q and token_q else None

            if not member_entry:
                st.warning("No matching member token found. You may still fill the form, but editing later will require this exact token+email link.")

            # check for existing response for this token
            resp_file = RESPONSES_DIR / f"{form_id_q}_{token_q or 'anon'}.json" if token_q else None
            existing_resp = load_json(resp_file) if resp_file and resp_file.exists() else None

            with st.form("response_form"):
                # render fields
                responses = {}
                for col in template["columns"]:
                    name = col["name"]
                    if col["type"] == "dropdown":
                        options = col.get("options", [])
                        default = existing_resp.get(name) if existing_resp else None
                        responses[name] = st.selectbox(name, [""] + options, index=(options.index(default)+1) if default in options else 0)
                    else:
                        default = existing_resp.get(name) if existing_resp else ""
                        responses[name] = st.text_input(name, value=default or "")
                submitted = st.form_submit_button("Submit Response")
                if submitted:
                    # save response
                    record = {
                        "form_id": form_id_q,
                        "token": token_q or uuid.uuid4().hex,
                        "email": email_q,
                        "data": responses,
                        "submitted_at": datetime.utcnow().isoformat()
                    }
                    # write to RESPONSES_DIR
                    token_val = record["token"]
                    save_json(RESPONSES_DIR / f"{form_id_q}_{token_val}.json", record)
                    st.success("Response saved. Thank you!")

                    # send confirmation emails
                    # to respondent
                    if email_q:
                        body_user = f"Thank you for submitting the form '{template.get('title')}'. Your responses:\n\n"
                        for k,v in responses.items():
                            body_user += f"{k}: {v}\n"
                        body_user += f"\nYou can edit your response using the same link."
                        if sender_email and sender_pass:
                            ok, err = send_email(smtp_server, smtp_port, sender_email, sender_pass, email_q, f"Confirmation: {template.get('title')}", body_user)
                            if ok:
                                st.info("Confirmation email sent to respondent.")
                            else:
                                st.warning(f"Could not send confirmation email: {err}")
                        else:
                            st.info("No sender credentials provided; confirmation email not sent.")

                    # to form owner/sender
                    if sender_email:
                        owner_body = f"New response received for form '{template.get('title')}' from {email_q or 'Unknown'}:\n\n"
                        for k,v in responses.items():
                            owner_body += f"{k}: {v}\n"
                        if sender_email and sender_pass:
                            ok, err = send_email(smtp_server, smtp_port, sender_email, sender_pass, sender_email, f"New response: {template.get('title')}", owner_body)
                            if ok:
                                st.info("Notification email sent to owner.")
                            else:
                                st.warning(f"Owner notification not sent: {err}")

# ----------------------------
# Tab 3: Dashboard / Responses
# ----------------------------
with tab_dash:
    st.header("3) Dashboard - Forms, Members & Responses")
    # list templates
    templates = list(TEMPLATES_DIR.glob("*.json"))
    t_select = st.selectbox("Select Form", ["-- choose --"] + [t.stem for t in templates])
    if t_select and t_select != "-- choose --":
        form_id = t_select
        tpl = load_json(TEMPLATES_DIR / f"{form_id}.json")
        st.subheader(tpl.get("title"))
        st.write("Description:", tpl.get("description"))
        st.markdown("**Columns**")
        st.table(pd.DataFrame(tpl.get("columns",[])))

        # members
        members_path = MEMBERS_DIR / f"{form_id}_members.json"
        members = load_json(members_path) or []
        st.markdown("**Members & Tokens**")
        if members:
            st.dataframe(pd.DataFrame(members))
        else:
            st.info("No members recorded for this form (maybe not emailed yet).")

        # responses
        st.markdown("**Responses**")
        resp_files = list(RESPONSES_DIR.glob(f"{form_id}_*.json"))
        if resp_files:
            all_resps = []
            for rf in resp_files:
                r = load_json(rf)
                flattened = {"token": r.get("token"), "email": r.get("email"), "submitted_at": r.get("submitted_at")}
                flattened.update(r.get("data", {}))
                all_resps.append(flattened)
            df_resps = pd.DataFrame(all_resps)
            st.dataframe(df_resps)

            # allow export
            if st.button("Export responses to Excel"):
                out_df = df_resps.copy()
                out_path = BASE / f"{form_id}_responses_export.xlsx"
                out_df.to_excel(out_path, index=False)
                st.success(f"Exported to `{out_path}`")

            # allow select response to edit
            sel_token = st.selectbox("Select token to edit", ["-- select --"] + [r.get("token") for r in all_resps])
            if sel_token and sel_token != "-- select --":
                target_file = RESPONSES_DIR / f"{form_id}_{sel_token}.json"
                target = load_json(target_file)
                st.markdown("Edit response fields below and Save.")
                if target:
                    edit_form = st.form("edit_response_form")
                    new_data = {}
                    for col in tpl.get("columns", []):
                        if col.get("type") == "dropdown":
                            new_data[col["name"]] = edit_form.selectbox(col["name"], [""] + col.get("options", []), index=0 if not target["data"].get(col["name"]) else (col.get("options",[]).index(target["data"].get(col["name"]))+1))
                        else:
                            new_data[col["name"]] = edit_form.text_input(col["name"], value=target["data"].get(col["name"], ""))
                    save_edit = edit_form.form_submit_button("Save edited response")
                    if save_edit:
                        target["data"] = new_data
                        target["edited_at"] = datetime.utcnow().isoformat()
                        save_json(target_file, target)
                        st.success("Saved edited response.")
                        st.experimental_rerun()
        else:
            st.info("No responses yet.")

    st.markdown("---")
    st.subheader("Sent Emails Log")
    sent_log = load_json(SENT_EMAILS_LOG) or []
    if sent_log:
        st.dataframe(pd.DataFrame(sent_log))
    else:
        st.info("No emails logged yet.")
