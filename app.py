# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import urllib.parse
import json
import uuid
import os
from datetime import datetime
import streamlit.components.v1 as components

# ---------------------------
# Config / storage files
# ---------------------------
META_FILE = "forms_meta.json"  # maps form_id -> metadata (columns, dropdowns, created_at, name)
STORAGE_DIR = "form_storage"   # where submissions and member status are saved
os.makedirs(STORAGE_DIR, exist_ok=True)

st.set_page_config(page_title="Auto Form Link + WhatsApp + Tracker", page_icon="📄", layout="centered")
st.title("📄 Auto Form Link Generator + WhatsApp (API-free)")

# ---------------------------
# Helpers
# ---------------------------
def load_meta():
    if os.path.exists(META_FILE):
        with open(META_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_meta(meta):
    with open(META_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

def detect_header_row(excel_file):
    df_raw = pd.read_excel(excel_file, header=None)
    header_row_index = None
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        if row.notna().sum() > 2:
            header_row_index = i
            break
    return header_row_index

def detect_dropdowns(form_file, df_columns):
    form_file.seek(0)
    wb = load_workbook(form_file, data_only=True)
    ws = wb.active
    dropdown_dict = {}
    if ws.data_validations is not None:
        for dv in ws.data_validations.dataValidation:
            try:
                if dv.type == "list" and dv.formula1:
                    formula = str(dv.formula1).strip('"')
                    if "," in formula:
                        values = [v.strip() for v in formula.split(",")]
                    else:
                        # try to resolve range references like Sheet!$A$1:$A$5
                        values = []
                        try:
                            rng = formula.split("!")[-1].replace("$", "")
                            cells = rng.split(":")
                            if len(cells) == 2:
                                start = cells[0]
                                end = cells[1]
                                col_letters = re.match(r"([A-Za-z]+)", start).group(1)
                                start_row = int(re.match(r"[A-Za-z]+([0-9]+)", start).group(1))
                                end_row = int(re.match(r"[A-Za-z]+([0-9]+)", end).group(1))
                                col_index = column_index_from_string(col_letters)
                                for r in range(start_row, end_row + 1):
                                    cell_val = ws.cell(row=r, column=col_index).value
                                    if cell_val is not None:
                                        values.append(str(cell_val))
                        except Exception:
                            values = []
                    for cell_range in dv.cells:
                        try:
                            if hasattr(cell_range, "min_col"):
                                col_index = cell_range.min_col - 1
                            else:
                                s = str(cell_range).split(":")[0]
                                match = re.match(r"([A-Za-z]+)", s)
                                if not match:
                                    continue
                                col_letters = match.group(1)
                                col_index = column_index_from_string(col_letters) - 1
                            if 0 <= col_index < len(df_columns):
                                dropdown_dict[df_columns[col_index]] = values
                        except Exception:
                            continue
            except Exception:
                continue
    return dropdown_dict

def make_whatsapp_url(number, message):
    encoded = urllib.parse.quote(message)
    return f"https://wa.me/{number}?text={encoded}"

def normalize_phone(ph):
    if pd.isna(ph):
        return ""
    s = str(ph).strip()
    s = s.replace("+", "").replace(" ", "").replace("-", "")
    return s

def save_submissions(form_id, submissions_list):
    # submissions_list: list of dicts
    df = pd.DataFrame(submissions_list)
    path = os.path.join(STORAGE_DIR, f"submissions_{form_id}.xlsx")
    df.to_excel(path, index=False)
    return path

def load_submissions(form_id):
    path = os.path.join(STORAGE_DIR, f"submissions_{form_id}.xlsx")
    if os.path.exists(path):
        return pd.read_excel(path)
    return pd.DataFrame()

def save_members_status(form_id, df_status):
    path = os.path.join(STORAGE_DIR, f"members_status_{form_id}.xlsx")
    df_status.to_excel(path, index=False)
    return path

def load_members_status(form_id):
    path = os.path.join(STORAGE_DIR, f"members_status_{form_id}.xlsx")
    if os.path.exists(path):
        return pd.read_excel(path)
    return pd.DataFrame(columns=["Name", "Whatsapp", "Status", "LastSubmitted"])

# ---------------------------
# Determine mode: Member view (form_id in URL) or Admin
# ---------------------------
query_params = st.experimental_get_query_params()
meta = load_meta()
form_id_in_url = query_params.get("form_id", [None])[0]

# ---------------------------
# MEMBER VIEW: if form_id provided -> show that form to user
# ---------------------------
if form_id_in_url:
    st.info(f"Form link detected: form_id = `{form_id_in_url}`")
    if form_id_in_url not in meta:
        st.error("This form does not exist or has expired. Contact the admin.")
    else:
        fm = meta[form_id_in_url]
        st.header(f"🔷 {fm.get('form_name','Form')}")

        # Load members status and submissions for this form
        ms = load_members_status(form_id_in_url)
        submissions_df = load_submissions(form_id_in_url)

        # member selection
        if not ms.empty:
            member_names = ms["Name"].astype(str).tolist()
            selected_name = st.selectbox("Your Name (select from the list):", options=["-- I'm not listed --"] + member_names)
            if selected_name == "-- I'm not listed --":
                submit_name = st.text_input("Type your full name as the admin has it:")
            else:
                submit_name = selected_name
        else:
            submit_name = st.text_input("Your Name:")

        st.write("Please fill the form fields below and press Submit.")
        # show form fields based on stored columns & dropdowns
        stored_cols = fm["columns"]
        stored_dropdowns = fm.get("dropdowns", {})

        form_values = {}
        for col in stored_cols:
            # skip 'Name' column inside form fields to avoid duplication
            if col.strip().lower() == "name":
                continue
            if col in stored_dropdowns and isinstance(stored_dropdowns[col], list) and stored_dropdowns[col]:
                form_values[col] = st.selectbox(col, stored_dropdowns[col], key=f"f_{col}")
            else:
                form_values[col] = st.text_input(col, key=f"f_{col}")

        if st.button("✅ Submit Form"):
            if not submit_name or str(submit_name).strip() == "":
                st.error("Please provide your name so we can record you.")
            else:
                # append submission to storage (server-side)
                submission = {"_SubmittedBy": submit_name, "_SubmittedAt": datetime.utcnow().isoformat()}
                submission.update(form_values)
                # load old submissions, append, save
                existing = load_submissions(form_id_in_url)
                if existing.empty:
                    new_df = pd.DataFrame([submission])
                else:
                    new_df = pd.concat([existing, pd.DataFrame([submission])], ignore_index=True)
                new_path = os.path.join(STORAGE_DIR, f"submissions_{form_id_in_url}.xlsx")
                new_df.to_excel(new_path, index=False)
                st.success("🎉 Your response has been recorded. Thank you!")

                # update members status if member found
                ms_local = load_members_status(form_id_in_url)
                if not ms_local.empty:
                    mask = ms_local["Name"].astype(str).str.strip().str.lower() == str(submit_name).strip().lower()
                    if mask.any():
                        idx = ms_local[mask].index[0]
                        ms_local.at[idx, "Status"] = "✅ Filled"
                        ms_local.at[idx, "LastSubmitted"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
                    else:
                        # append new row if name not found
                        ms_local = pd.concat([ms_local, pd.DataFrame([{
                            "Name": submit_name, "Whatsapp": "", "Status": "✅ Filled",
                            "LastSubmitted": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
                        }])], ignore_index=True)
                    save_members_status(form_id_in_url, ms_local)
                else:
                    # create new members_status file
                    new_ms = pd.DataFrame([{
                        "Name": submit_name, "Whatsapp": "", "Status": "✅ Filled",
                        "LastSubmitted": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
                    }])
                    save_members_status(form_id_in_url, new_ms)

        st.markdown("---")
        st.subheader("Form Info & Progress")
        # show progress if members list exists
        ms_display = load_members_status(form_id_in_url)
        if not ms_display.empty:
            total = len(ms_display)
            filled = (ms_display["Status"] == "✅ Filled").sum()
            st.progress(filled / total if total > 0 else 0)
            st.write(f"✅ {filled} of {total} members have submitted the form")
            st.dataframe(ms_display.reset_index(drop=True))
        else:
            st.write("No members information available for this form.")

        st.markdown("You can close this tab after submission. If you think your submission didn't record, contact admin.")

# ---------------------------
# ADMIN VIEW (no form_id in url)
# ---------------------------
else:
    st.header("Admin Dashboard — Create Auto Form Link & Send WhatsApp Links")

    # Upload members
    st.subheader("1) Upload Members List (Excel)")
    st.write("Members Excel must have columns: `Name` and `Whatsapp` (e.g. 923001234567).")
    members_file = st.file_uploader("Upload Members Excel", type=["xlsx"], key="admin_members")
    members_df = None
    if members_file:
        try:
            members_df = pd.read_excel(members_file)
            members_df.columns = [str(c).strip().title() for c in members_df.columns]
            if "Name" not in members_df.columns or "Whatsapp" not in members_df.columns:
                st.error("Members file must contain 'Name' and 'Whatsapp' columns.")
                members_df = None
            else:
                members_df["Whatsapp"] = members_df["Whatsapp"].apply(normalize_phone)
                members_df["Status"] = "❌ Pending"
                members_df["LastSubmitted"] = ""
                st.success("Members loaded.")
                st.dataframe(members_df[["Name", "Whatsapp"]].reset_index(drop=True))
        except Exception as e:
            st.error(f"Error reading members file: {e}")
            members_df = None

    # Upload form template
    st.subheader("2) Upload Form Template (Excel)")
    st.write("Upload the Excel file used as form template. Columns and dropdowns will be detected.")
    form_file = st.file_uploader("Upload Form Excel", type=["xlsx"], key="admin_form")
    if form_file:
        try:
            header_idx = detect_header_row(form_file)
            form_df = pd.read_excel(form_file, header=header_idx)
            form_df.columns = [str(c).strip().replace("_", " ").title() for c in form_df.columns if pd.notna(c)]
            st.write("Detected columns:", list(form_df.columns))
            dropdowns = detect_dropdowns(form_file, list(form_df.columns))
            if dropdowns:
                st.write("Detected dropdowns:")
                st.table(pd.DataFrame([{"Column": c, "Options": ", ".join(v)} for c, v in dropdowns.items()]))
        except Exception as e:
            st.error(f"Error processing form file: {e}")
            form_df = None
            dropdowns = {}

        # Create form link (auto) once admin confirms
        st.markdown("### Create form link from this uploaded template")
        form_name = st.text_input("Optional: Give this form a name (e.g. 'Retirement Form - Nov 2025')", value=f"Form_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}")
        if st.button("🔖 Generate Form Link and Save (auto)"):
            if form_df is None:
                st.error("Upload a valid form Excel first.")
            else:
                # generate unique form_id
                new_form_id = str(uuid.uuid4())[:12]
                meta = load_meta()
                meta[new_form_id] = {
                    "form_name": form_name,
                    "created_at": datetime.utcnow().isoformat(),
                    "columns": list(form_df.columns),
                    "dropdowns": dropdowns
                }
                save_meta(meta)
                # save members status for this form
                if members_df is not None:
                    ms = members_df[["Name", "Whatsapp", "Status", "LastSubmitted"]].copy()
                else:
                    ms = pd.DataFrame(columns=["Name", "Whatsapp", "Status", "LastSubmitted"])
                save_members_status(new_form_id, ms)
                # create empty submissions file
                save_submissions(new_form_id, [])
                st.success(f"Form created with id: {new_form_id}")

                # create link automatically (app must be on public URL). We'll build relative link for admin to copy.
                # Try to infer base URL from request (not always available) — we will instruct admin to copy the link below.
                base = st.experimental_get_query_params().get("base_url", [""])[0]
                # Best approach: ask admin to paste the public app base URL OR attempt to build from request
                st.info("Copy the generated link below and use it in WhatsApp messages.")
                # use generic pattern; admin should replace with their public URL if needed
                suggested_link = f"?form_id={new_form_id}"
                st.code(f"Streamlit app URL + {suggested_link}\n\nExample final link (if your app is at https://myapp.streamlit.app):\nhttps://myapp.streamlit.app/?form_id={new_form_id}")

    # Show existing forms (meta)
    st.subheader("Existing Forms")
    meta = load_meta()
    if meta:
        rows = []
        for fid, info in meta.items():
            created = info.get("created_at", "")
            name = info.get("form_name", "")
            cols = ", ".join(info.get("columns", []))[:120]
            rows.append({"Form ID": fid, "Name": name, "Created": created, "Columns": cols})
        st.table(pd.DataFrame(rows))
    else:
        st.info("No forms created yet.")

    # If admin chooses a form id, show admin controls for that form (send links, view progress, downloads)
    st.markdown("---")
    st.subheader("Admin: Manage a Form (Send links, View status)")
    chosen_form_id = st.text_input("Enter a Form ID to manage (copy from 'Existing Forms'):")
    if chosen_form_id:
        meta = load_meta()
        if chosen_form_id not in meta:
            st.error("Unknown Form ID.")
        else:
            info = meta[chosen_form_id]
            st.write(f"Form: **{info.get('form_name','-')}** — Created: {info.get('created_at')}")
            ms = load_members_status(chosen_form_id)
            subs = load_submissions(chosen_form_id)

            # WhatsApp message template (auto include link)
            st.markdown("### WhatsApp Message Template (will include form link automatically)")
            default_tpl = "Hello {name}! Please fill your form here: {link}"
            tpl = st.text_area("Message template (use {name} and {link} placeholders):", value=default_tpl, height=80)
            if "{link}" not in tpl:
                st.error("Please include {link} in message template.")
            else:
                # Build links
                wa_links = []
                rows = []
                for _, r in ms.iterrows():
                    name = r["Name"]
                    phone = normalize_phone(r["Whatsapp"])
                    link_for_member = f"?form_id={chosen_form_id}"
                    # admin needs to prepend their app base URL to this (we show example)
                    message = tpl.replace("{name}", name).replace("{link}", f"{st.experimental_get_query_params().get('app_base','[PASTE_APP_BASE_URL]')}{link_for_member}")
                    wa_url = make_whatsapp_url(phone, message)
                    wa_links.append(wa_url)
                    rows.append({"Name": name, "Whatsapp": phone, "Send Link": f"[Send]({wa_url})", "Status": r.get("Status", "❌ Pending")})
                if rows:
                    st.markdown(pd.DataFrame(rows).to_markdown(index=False), unsafe_allow_html=True)
                    if st.button("🚀 Open All WhatsApp (may be blocked by popups)"):
                        js = "<script>const links = " + str(wa_links).replace("'", '"') + ";\nfor(let i=0;i<links.length;i++){ window.open(links[i], '_blank'); }</script>"
                        components.html(js, height=10)
                else:
                    st.write("Members list is empty for this form.")

            # Show progress & downloads
            st.markdown("### Progress & Downloads")
            if not ms.empty:
                total = len(ms)
                filled = (ms["Status"] == "✅ Filled").sum()
                st.progress(filled / total if total > 0 else 0)
                st.write(f"✅ {filled} of {total} members have submitted")
                st.dataframe(ms.reset_index(drop=True))
            else:
                st.write("No members data available (upload members when creating form).")

            if not subs.empty:
                st.write("Submissions (latest first):")
                st.dataframe(subs.iloc[::-1].reset_index(drop=True))
            else:
                st.write("No submissions for this form yet.")

            # Save / Export
            if st.button("⬇️ Download Submissions Excel"):
                path = os.path.join(STORAGE_DIR, f"submissions_{chosen_form_id}.xlsx")
                if os.path.exists(path):
                    with open(path, "rb") as f:
                        st.download_button("Download", f, file_name=f"submissions_{chosen_form_id}.xlsx")
                else:
                    st.error("No submission file found.")

            if st.button("⬇️ Download Members Status Excel"):
                path2 = os.path.join(STORAGE_DIR, f"members_status_{chosen_form_id}.xlsx")
                if os.path.exists(path2):
                    with open(path2, "rb") as f:
                        st.download_button("Download", f, file_name=f"members_status_{chosen_form_id}.xlsx")
                else:
                    st.error("No members status file found.")

    st.markdown("---")
    st.caption("Notes: 1) This app creates an internal form_id for each uploaded form and stores submissions & statuses server-side in the 'form_storage' folder. 2) WhatsApp messages are prepared using wa.me links — sending requires clicking 'Send' in WhatsApp. 3) The generated link shown is a relative pattern (?form_id=...). Replace with your public Streamlit URL prefix when sending to members (example provided when generating the form).")

