# app.py
import streamlit as st
import pandas as pd
import os
import json
import re
import uuid
import urllib.parse
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import streamlit.components.v1 as components

# ---------------------------
# Config
# ---------------------------
st.set_page_config(page_title="Excel Form + WhatsApp (Admin ↔ Form)", layout="centered")
st.title("📄 Excel Form Generator + WhatsApp Links + Persistent Tracking")

DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")

# ---------------------------
# Helpers
# ---------------------------
def load_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_meta(meta):
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

def normalize_phone(ph):
    if pd.isna(ph): return ""
    s = str(ph).strip()
    s = s.replace("+", "").replace(" ", "").replace("-", "")
    # If local 0-prefixed Pakistan number, convert to 92...
    if s.startswith("0") and not s.startswith("92"):
        s = "92" + s[1:]
    return s

def detect_header_row(excel_file):
    df_raw = pd.read_excel(excel_file, header=None)
    for i in range(len(df_raw)):
        if df_raw.iloc[i].notna().sum() > 2:
            return i
    return 0

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
                    # inline list like "A,B,C"
                    if "," in formula:
                        options = [x.strip() for x in formula.split(",")]
                    else:
                        # try to resolve reference like Sheet!$A$1:$A$5
                        options = []
                        try:
                            rng = formula.split("!")[-1].replace("$", "")
                            a,b = rng.split(":")
                            col_letters = re.match(r"([A-Za-z]+)", a).group(1)
                            start_row = int(re.match(r"[A-Za-z]+([0-9]+)", a).group(1))
                            end_row = int(re.match(r"[A-Za-z]+([0-9]+)", b).group(1))
                            col_idx = column_index_from_string(col_letters)
                            for r in range(start_row, end_row+1):
                                v = ws.cell(row=r, column=col_idx).value
                                if v is not None:
                                    options.append(str(v))
                        except Exception:
                            options = []
                    # map affected columns
                    for cell_range in dv.cells:
                        try:
                            if hasattr(cell_range, "min_col"):
                                cidx = cell_range.min_col - 1
                            else:
                                s = str(cell_range).split(":")[0]
                                m = re.match(r"([A-Za-z]+)", s)
                                if not m:
                                    continue
                                col_letters = m.group(1)
                                cidx = column_index_from_string(col_letters) - 1
                            if 0 <= cidx < len(df_columns):
                                dropdowns[df_columns[cidx]] = options
                        except Exception:
                            continue
            except Exception:
                continue
    return dropdowns

def whatsapp_url(number, message):
    return f"https://wa.me/{number}?text={urllib.parse.quote(message)}"

def read_saved_submissions(form_id):
    path = os.path.join(DATA_DIR, f"submissions_{form_id}.xlsx")
    if os.path.exists(path):
        return pd.read_excel(path)
    return pd.DataFrame()

def save_submissions_df(form_id, df):
    path = os.path.join(DATA_DIR, f"submissions_{form_id}.xlsx")
    df.to_excel(path, index=False)
    return path

def read_members_status(form_id):
    path = os.path.join(DATA_DIR, f"members_{form_id}.xlsx")
    if os.path.exists(path):
        return pd.read_excel(path)
    return pd.DataFrame(columns=["Name","Whatsapp","Status","LastSubmitted"])

def save_members_status(form_id, df):
    path = os.path.join(DATA_DIR, f"members_{form_id}.xlsx")
    df.to_excel(path, index=False)
    return path

# ---------------------------
# URL params: mode & form_id
# ---------------------------
params = st.experimental_get_query_params()
mode = params.get("mode", ["admin"])[0]  # default admin
param_form_id = params.get("form_id", [None])[0]

meta = load_meta()
latest_form = meta.get("latest_form_id")

# If no explicit form_id, use latest when in form mode
form_id = param_form_id if param_form_id else latest_form

# ---------------------------
# MEMBER VIEW (public form)
# ---------------------------
if mode == "form":
    st.markdown("### 🧾 Public Form")
    if not form_id or form_id not in meta:
        st.error("No active form found. Contact admin.")
    else:
        info = meta[form_id]
        st.write(f"**Form name:** {info.get('form_name')}")
        cols = info.get("columns", [])
        dropdowns = info.get("dropdowns", {})

        # Try load members for name select (optional)
        members_df = read_members_status(form_id)
        if not members_df.empty:
            name_options = ["-- I'm not listed --"] + members_df["Name"].astype(str).tolist()
            selected = st.selectbox("Select your name (choose from list) or pick 'I'm not listed' and type below:", options=name_options)
            if selected == "-- I'm not listed --":
                typed_name = st.text_input("Type your full name:")
                submit_name = typed_name.strip()
            else:
                submit_name = selected
        else:
            submit_name = st.text_input("Your Name:").strip()

        st.write("Please fill the form below and press Submit.")
        values = {}
        for c in cols:
            if c.strip().lower() == "name":
                continue
            if c in dropdowns and isinstance(dropdowns[c], list) and dropdowns[c]:
                values[c] = st.selectbox(c, dropdowns[c], key=f"f_{c}")
            else:
                values[c] = st.text_input(c, key=f"f_{c}")

        if st.button("✅ Submit"):
            if not submit_name:
                st.error("Please provide your name so we can record who submitted.")
            else:
                # build new submission row
                row = {"Name": submit_name, "SubmittedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                row.update(values)
                subs = read_saved_submissions(form_id)
                subs = pd.concat([subs, pd.DataFrame([row])], ignore_index=True) if not subs.empty else pd.DataFrame([row])
                save_submissions_df(form_id, subs)
                # update member status if exists
                ms = read_members_status(form_id)
                if not ms.empty:
                    mask = ms["Name"].astype(str).str.strip().str.lower() == submit_name.strip().lower()
                    if mask.any():
                        idx = ms[mask].index[0]
                        ms.at[idx, "Status"] = "✅ Filled"
                        ms.at[idx, "LastSubmitted"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        # add new row
                        ms = pd.concat([ms, pd.DataFrame([{
                            "Name": submit_name, "Whatsapp": "", "Status": "✅ Filled", "LastSubmitted": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }])], ignore_index=True)
                else:
                    ms = pd.DataFrame([{
                        "Name": submit_name, "Whatsapp": "", "Status": "✅ Filled", "LastSubmitted": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }])
                save_members_status(form_id, ms)
                st.success("🎉 Thank you — your response has been recorded.")

        st.markdown("---")
        # show progress to submitter (optional)
        ms_display = read_members_status(form_id)
        if not ms_display.empty:
            total = len(ms_display)
            filled = (ms_display["Status"] == "✅ Filled").sum()
            st.write(f"Progress: {filled} / {total} submitted")
        else:
            st.write("No member list available for this form.")

# ---------------------------
# ADMIN VIEW
# ---------------------------
else:
    st.header("🧑‍💼 Admin Dashboard — Create & Manage Form")

    st.markdown("#### 1) Upload Members Excel (must have `Name` and `Whatsapp` columns)")
    members_file = st.file_uploader("Upload members file (Excel)", type=["xlsx"], key="members_upload")

    st.markdown("#### 2) Upload Form Template Excel (columns become form fields; dropdowns auto-detected)")
    form_file = st.file_uploader("Upload form template (Excel)", type=["xlsx"], key="form_upload")

    # Show existing forms quick list
    st.markdown("---")
    st.subheader("Existing forms (created earlier)")
    if meta:
        rows = []
        for fid, info in meta.get("forms", {}).items() if "forms" in meta else []:
            pass
    # Show table of existing forms
    all_meta = load_meta()
    forms = all_meta.get("forms", {}) if all_meta else {}
    if forms:
        df_forms = pd.DataFrame([{
            "Form ID": fid,
            "Form Name": info.get("form_name"),
            "Created": info.get("created_at")[:19],
            "Fields": ", ".join(info.get("columns", [])[:6])
        } for fid, info in forms.items()])
        st.dataframe(df_forms)
    else:
        st.info("No forms created yet. Create one by uploading members and form template below.")

    # Handle uploads and generate
    if members_file is not None and form_file is not None:
        try:
            members_df = pd.read_excel(members_file)
            members_df.columns = [str(c).strip().title() for c in members_df.columns]
            if not {"Name", "Whatsapp"}.issubset(set(members_df.columns)):
                st.error("Members file must contain 'Name' and 'Whatsapp' columns.")
            else:
                # normalize phones and initial status
                members_df["Whatsapp"] = members_df["Whatsapp"].apply(normalize_phone)
                members_df["Status"] = "❌ Pending"
                members_df["LastSubmitted"] = ""

                # Detect header row & read form fields
                hdr = detect_header_row(form_file)
                form_df = pd.read_excel(form_file, header=hdr)
                form_df.columns = [str(c).strip().replace("_"," ").title() for c in form_df.columns if pd.notna(c)]
                st.success(f"Detected form columns: {len(form_df.columns)}")
                st.write(form_df.columns.tolist())

                # Detect dropdowns
                dropdowns = detect_dropdowns(form_file, list(form_df.columns))
                if dropdowns:
                    st.info("Detected dropdowns:")
                    st.table(pd.DataFrame([{"Field": k, "Options": ", ".join(v)} for k,v in dropdowns.items()]))

                # Inputs for creating form
                form_name = st.text_input("Form name (optional)", value=f"My Form {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                base_url = st.text_input("Your public Streamlit app URL (example: https://yourapp.streamlit.app) — required to build shareable link")

                if st.button("🔖 Generate Form Link & Save"):
                    if not base_url:
                        st.error("Please enter your public app base URL so shareable links include the full address.")
                    else:
                        # create form_id and save meta + files
                        new_form_id = str(uuid.uuid4())[:10]
                        forms_meta = all_meta.get("forms", {})
                        forms_meta[new_form_id] = {
                            "form_name": form_name,
                            "created_at": datetime.now().isoformat(),
                            "columns": list(form_df.columns),
                            "dropdowns": dropdowns
                        }
                        all_meta["forms"] = forms_meta
                        all_meta["latest_form_id"] = new_form_id
                        save_meta(all_meta)

                        # save members status file for this form
                        save_members_status(new_form_id, members_df[["Name","Whatsapp","Status","LastSubmitted"]])

                        # create empty submissions file
                        save_submissions_path = os.path.join(DATA_DIR, f"submissions_{new_form_id}.xlsx")
                        pd.DataFrame(columns=["Name","SubmittedAt"] + list(form_df.columns)).to_excel(save_submissions_path, index=False)

                        # show link
                        share_link = base_url.rstrip("/") + f"/?mode=form&form_id={new_form_id}"
                        st.success("Form created and saved.")
                        st.info("Share this link with your members (same link for all):")
                        st.code(share_link)

                        # Prepare WhatsApp links table for admin to click / download
                        wa_rows = []
                        for _, r in members_df.iterrows():
                            name = r["Name"]
                            phone = r["Whatsapp"]
                            msg = f"Hello {name}! Please fill your form here: {share_link}"
                            wa = whatsapp_url(phone, msg)
                            wa_rows.append({"Name": name, "Whatsapp": phone, "WhatsApp Link": wa})
                        wa_df = pd.DataFrame(wa_rows)
                        st.subheader("WhatsApp Links (click to open chat with message ready)")
                        st.markdown(wa_df.to_markdown(index=False), unsafe_allow_html=True)

                        # Open all button (may be popup-blocked)
                        if st.button("🚀 Open All WhatsApp Links (might be blocked by popup blocker)"):
                            urls = wa_df["WhatsApp Link"].tolist()
                            js = "<script>\n"
                            js += "const urls = " + str(urls).replace("'", '"') + ";\n"
                            js += "for (let i=0;i<urls.length;i++){ window.open(urls[i], '_blank'); }\n"
                            js += "</script>"
                            components.html(js, height=10)

                        # allow admin to download WA links excel
                        buf = BytesIO()
                        wa_df.to_excel(buf, index=False)
                        buf.seek(0)
                        st.download_button("⬇️ Download WhatsApp Links (Excel)", data=buf, file_name=f"wa_links_{new_form_id}.xlsx")

        except Exception as e:
            st.error(f"Error processing uploaded files: {e}")

    st.markdown("---")
    # Admin choose a form to manage (or use latest)
    st.subheader("Manage existing form (view progress, submissions, download)")
    forms_all = all_meta.get("forms", {}) if all_meta else {}
    if forms_all:
        form_ids = list(forms_all.keys())
        chosen = st.selectbox("Choose form to manage (or choose latest):", options=["-- latest --"] + form_ids)
        if chosen == "-- latest --":
            chosen_id = all_meta.get("latest_form_id")
        else:
            chosen_id = chosen

        if chosen_id:
            info = forms_all.get(chosen_id, {})
            st.write(f"**Form:** {info.get('form_name')}  (ID: {chosen_id})")
            ms = read_members_status(chosen_id)
            subs = read_saved_submissions(chosen_id)
            if not ms.empty:
                total = len(ms)
                filled = (ms["Status"] == "✅ Filled").sum()
                st.progress(filled / total if total>0 else 0)
                st.write(f"✅ {filled} of {total} members have submitted")
                st.subheader("Members Status")
                st.dataframe(ms)
            else:
                st.info("No members file saved for this form (maybe none uploaded).")

            st.subheader("Submissions")
            if not subs.empty:
                st.dataframe(subs.iloc[::-1].reset_index(drop=True))
                # download submissions
                buf2 = BytesIO()
                subs.to_excel(buf2, index=False)
                buf2.seek(0)
                st.download_button("⬇️ Download Submissions (Excel)", data=buf2, file_name=f"submissions_{chosen_id}.xlsx")
            else:
                st.write("No submissions yet for this form.")
    else:
        st.info("No saved forms yet. Create one above.")

    st.markdown("---")
    st.caption("Notes: 1) This app does NOT send WhatsApp automatically — it prepares wa.me links for manual sending. 2) 'Open All' may be blocked by browser popup blockers. 3) All data saved to the data_store folder.")

