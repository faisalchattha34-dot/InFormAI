# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import os
import json
import urllib.parse
from datetime import datetime

# -------------------------
# Config / storage
# -------------------------
st.set_page_config(page_title="Excel Form (Admin / Form)", page_icon="📄", layout="centered")
st.title("📄 Excel Smart Form — Admin / Public Form (mode=admin | mode=form)")

DATA_DIR = "app_data"
os.makedirs(DATA_DIR, exist_ok=True)

META_PATH = os.path.join(DATA_DIR, "current_form_meta.json")      # stores current form columns & dropdowns
SUBMISSIONS_PATH = os.path.join(DATA_DIR, "submissions.xlsx")    # persistent submissions
MEMBERS_PATH = os.path.join(DATA_DIR, "members_status.xlsx")     # persistent members + status

# -------------------------
# Helpers
# -------------------------
def detect_header_row(excel_file):
    df_raw = pd.read_excel(excel_file, header=None)
    for i in range(len(df_raw)):
        if df_raw.iloc[i].notna().sum() > 1:
            return i
    return 0

def detect_dropdowns(file_like, df_columns):
    dropdowns = {}
    try:
        file_like.seek(0)
        wb = load_workbook(file_like, data_only=True)
        ws = wb.active
        if ws.data_validations is None:
            return {}
        for dv in ws.data_validations.dataValidation:
            try:
                if dv.type == "list" and dv.formula1:
                    formula = str(dv.formula1).strip('"')
                    values = []
                    if "," in formula:
                        values = [v.strip() for v in formula.split(",")]
                    else:
                        # range reference like Sheet!$A$1:$A$5 or $A$1:$A$5
                        try:
                            rng = formula.split("!")[-1].replace("$", "")
                            start, end = rng.split(":")
                            col_letters = re.match(r"([A-Za-z]+)", start).group(1)
                            start_row = int(re.match(r"[A-Za-z]+([0-9]+)", start).group(1))
                            end_row = int(re.match(r"[A-Za-z]+([0-9]+)", end).group(1))
                            col_idx = column_index_from_string(col_letters)
                            for r in range(start_row, end_row + 1):
                                v = ws.cell(row=r, column=col_idx).value
                                if v is not None:
                                    values.append(str(v))
                        except Exception:
                            values = []
                    # map values to affected columns
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
                                dropdowns[df_columns[col_index]] = values
                        except Exception:
                            continue
            except Exception:
                continue
    except Exception:
        return {}
    return dropdowns

def save_meta(meta: dict):
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

def load_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

def normalize_phone(ph):
    if pd.isna(ph):
        return ""
    s = str(ph).strip().replace("+", "").replace(" ", "").replace("-", "")
    # ensure begins with country code for Pakistan if needed (optional)
    if s.startswith("0"):
        s = "92" + s[1:]
    return s

def whatsapp_url(number, message):
    return f"https://wa.me/{number}?text={urllib.parse.quote(message)}"

def append_submission(row: dict):
    # ensure file exists or append
    if os.path.exists(SUBMISSIONS_PATH):
        df = pd.read_excel(SUBMISSIONS_PATH)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    else:
        df = pd.DataFrame([row])
    df.to_excel(SUBMISSIONS_PATH, index=False)

def load_submissions_df():
    if os.path.exists(SUBMISSIONS_PATH):
        return pd.read_excel(SUBMISSIONS_PATH)
    return pd.DataFrame()

def save_members_status(df_members: pd.DataFrame):
    df_members.to_excel(MEMBERS_PATH, index=False)

def load_members_status():
    if os.path.exists(MEMBERS_PATH):
        return pd.read_excel(MEMBERS_PATH)
    return pd.DataFrame(columns=["Name", "Whatsapp", "Status", "LastSubmitted"])

# -------------------------
# Routing by mode
# -------------------------
params = st.experimental_get_query_params()
mode = params.get("mode", ["admin"])[0]   # default admin
st.write(f"🔁 Running in mode: **{mode}**  — change URL param `?mode=form` or `?mode=admin`")

# -------------------------
# MEMBER (public) form mode
# -------------------------
if mode == "form":
    meta = load_meta()
    if not meta:
        st.error("No active form found. Contact admin.")
    else:
        st.header(f"📝 {meta.get('form_name','Form')}")
        cols = meta.get("columns", [])
        dropdowns = meta.get("dropdowns", {})

        # Member provides name (so admin can track)
        st.info("Please enter your name (exact as admin has it) so your submission is recorded against you.")
        name = st.text_input("Your Name")

        st.write("Please fill the form below:")
        form_values = {}
        for c in cols:
            # skip 'Name' column if it's present in template to avoid duplication
            if str(c).strip().lower() == "name":
                continue
            if c in dropdowns and isinstance(dropdowns[c], list) and dropdowns[c]:
                form_values[c] = st.selectbox(c, dropdowns[c], key=f"f_{c}")
            else:
                form_values[c] = st.text_input(c, key=f"f_{c}")

        if st.button("✅ Submit"):
            if not name or str(name).strip() == "":
                st.error("Please provide your name so the admin can match your submission.")
            else:
                entry = {"Name": name, **form_values, "SubmittedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                append_submission(entry)
                # update members status if present
                ms = load_members_status()
                if not ms.empty:
                    mask = ms["Name"].astype(str).str.strip().str.lower() == name.strip().lower()
                    if mask.any():
                        idx = ms[mask].index[0]
                        ms.at[idx, "Status"] = "✅ Filled"
                        ms.at[idx, "LastSubmitted"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        # append new member record
                        ms = pd.concat([ms, pd.DataFrame([{
                            "Name": name, "Whatsapp": "", "Status": "✅ Filled", "LastSubmitted": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }])], ignore_index=True)
                else:
                    ms = pd.DataFrame([{
                        "Name": name, "Whatsapp": "", "Status": "✅ Filled", "LastSubmitted": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }])
                save_members_status(ms)
                st.success("🎉 Your response has been recorded. Thank you!")

        st.markdown("---")
        st.subheader("Form Info")
        st.write(f"Form name: **{meta.get('form_name','Form')}**")
        st.write(f"Fields: {', '.join(cols)}")
        st.write("If you encounter any issue, contact admin.")

# -------------------------
# ADMIN mode
# -------------------------
else:
    st.header("🧑‍💼 Admin Dashboard — Upload Members & Form Template")

    st.markdown("**Step 1 — Upload Members Excel** (must have headers `Name` and `Whatsapp`)")
    members_file = st.file_uploader("Upload Members (Excel)", type=["xlsx"], key="members_uploader")
    st.markdown("**Step 2 — Upload Form Template Excel** (columns will become form fields and dropdowns are detected)")
    form_file = st.file_uploader("Upload Form Template (Excel)", type=["xlsx"], key="form_uploader")

    if members_file:
        try:
            members_df = pd.read_excel(members_file)
            # normalize header names
            members_df.columns = [str(c).strip().title() for c in members_df.columns]
            # try flexible matching for phone/name columns
            possible_name_cols = [c for c in members_df.columns if c.lower().strip() in ("name", "full name", "member name")]
            possible_phone_cols = [c for c in members_df.columns if c.lower().strip() in ("whatsapp", "phone", "mobile", "contact")]
            if not possible_name_cols or not possible_phone_cols:
                st.error("Members file must contain columns for Name and Whatsapp (or 'phone'/'mobile'). Please adjust headers.")
                members_df = None
            else:
                name_col = possible_name_cols[0]
                phone_col = possible_phone_cols[0]
                members_df = members_df.rename(columns={name_col: "Name", phone_col: "Whatsapp"})
                members_df["Whatsapp"] = members_df["Whatsapp"].astype(str).apply(normalize_phone)
                members_df["Status"] = "❌ Pending"
                members_df["LastSubmitted"] = ""
                st.success("Members loaded")
                st.dataframe(members_df[["Name", "Whatsapp", "Status"]].reset_index(drop=True))
        except Exception as e:
            st.error(f"Error reading members file: {e}")
            members_df = None

    if form_file:
        try:
            hdr = detect_header_row(form_file)
            form_df = pd.read_excel(form_file, header=hdr)
            form_df.columns = [str(c).strip().replace("_", " ").title() for c in form_df.columns if pd.notna(c)]
            st.success("Form columns detected")
            st.write("Columns:", list(form_df.columns))
            dropdowns = detect_dropdowns(form_file, list(form_df.columns))
            if dropdowns:
                st.info("Detected dropdowns:")
                st.table(pd.DataFrame([{"Column": c, "Options": ", ".join(v)} for c, v in dropdowns.items()]))
        except Exception as e:
            st.error(f"Error processing form file: {e}")
            form_df = None
            dropdowns = {}

    # When both uploaded, create/save meta and members_status
    if members_file and form_file:
        st.markdown("---")
        st.subheader("Create & Publish Form")
        form_name = st.text_input("Form Name (optional)", value=f"Form {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        base_url = st.text_input("Your public Streamlit app base URL (example: https://yourapp.streamlit.app)", help="Include protocol and domain, no trailing slash")
        if st.button("📤 Publish Form (generate form link & save members)"):
            if not base_url or not base_url.startswith("http"):
                st.error("Please enter your public app base URL (so generated link will work).")
            else:
                meta = {
                    "form_name": form_name,
                    "columns": list(form_df.columns),
                    "dropdowns": dropdowns,
                    "created_at": datetime.now().isoformat()
                }
                save_meta(meta)
                # save members status persistently
                save_members_status(members_df[["Name", "Whatsapp", "Status", "LastSubmitted"]])
                # ensure submissions file exists (create empty if not)
                if not os.path.exists(SUBMISSIONS_PATH):
                    pd.DataFrame(columns=["Name"] + [c for c in form_df.columns if str(c).strip().lower() != "name"] + ["SubmittedAt"]).to_excel(SUBMISSIONS_PATH, index=False)
                form_link = f"{base_url}?mode=form"
                st.success("Form published!")
                st.info(f"Share this link with members: {form_link}")
                # prepare whatsapp links for convenience
                msg_template = st.text_area("WhatsApp message template (use {name} and {link})", value="Hello {name}! Please fill this form: {link}")
                if "{link}" not in msg_template:
                    st.error("Include {link} placeholder in message template.")
                else:
                    wa_rows = []
                    for _, r in members_df.iterrows():
                        nm = r["Name"]
                        ph = r["Whatsapp"]
                        msg = msg_template.replace("{name}", nm).replace("{link}", form_link)
                        wa = whatsapp_url(ph, msg)
                        wa_rows.append({"Name": nm, "Whatsapp": ph, "WhatsApp Link": wa})
                    st.subheader("WhatsApp Links (click to open Chat)")
                    wa_df = pd.DataFrame(wa_rows)
                    # show clickable links in markdown table
                    md = wa_df.apply(lambda row: f"[Send to {row['Name']}]({row['WhatsApp Link']})", axis=1)
                    st.write(pd.concat([wa_df[["Name", "Whatsapp"]], md.rename("Send")], axis=1).to_markdown(index=False), unsafe_allow_html=True)
                    # allow download
                    buf = BytesIO()
                    wa_df.to_excel(buf, index=False)
                    buf.seek(0)
                    st.download_button("⬇️ Download WhatsApp links (Excel)", data=buf.getvalue(), file_name="wa_links.xlsx")

    # Admin area: show submissions and status
    st.markdown("---")
    st.subheader("Submissions & Progress (Admin View)")
    meta = load_meta()
    if not meta:
        st.info("No active form published yet.")
    else:
        st.write(f"Active Form: **{meta.get('form_name','Form')}** (created: {meta.get('created_at','-')})")
        # load members & submissions
        ms = load_members_status()
        subs = load_submissions_df()
        total = len(ms) if not ms.empty else 0
        filled = (ms["Status"] == "✅ Filled").sum() if not ms.empty else 0
        st.progress(filled / total if total > 0 else 0)
        st.write(f"✅ {filled} of {total} members have submitted")
        if not ms.empty:
            st.subheader("Members Status")
            st.dataframe(ms.reset_index(drop=True))
        else:
            st.write("No members uploaded for current form.")

        st.subheader("Submitted Data")
        if not subs.empty:
            st.dataframe(subs.iloc[::-1].reset_index(drop=True))
            # download submissions
            buf2 = By
