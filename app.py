import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import os, json, re, urllib.parse, uuid
from datetime import datetime
from io import BytesIO

# ==============================
# CONFIGURATION
# ==============================
st.set_page_config(page_title="Excel Form + WhatsApp Auto System", page_icon="📱", layout="centered")
st.title("📱 Excel Form Generator + WhatsApp Link + Live Tracking")

DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_FILE = os.path.join(DATA_DIR, "form_meta.json")

# ==============================
# HELPERS
# ==============================
def load_meta():
    if os.path.exists(META_FILE):
        with open(META_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_meta(meta):
    with open(META_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)

def normalize_phone(ph):
    if pd.isna(ph): return ""
    s = str(ph).replace("+", "").replace("-", "").replace(" ", "")
    if not s.startswith("92") and s.startswith("0"):
        s = "92" + s[1:]
    return s

def detect_header_row(excel):
    df_raw = pd.read_excel(excel, header=None)
    for i in range(len(df_raw)):
        if df_raw.iloc[i].notna().sum() > 2:
            return i
    return 0

def detect_dropdowns(file, df_cols):
    file.seek(0)
    wb = load_workbook(file, data_only=True)
    ws = wb.active
    dropdowns = {}
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1:
                vals = str(dv.formula1).strip('"')
                if "," in vals:
                    options = [v.strip() for v in vals.split(",")]
                    for cell_range in dv.cells:
                        try:
                            if hasattr(cell_range, "min_col"):
                                c = cell_range.min_col - 1
                            else:
                                m = re.match(r"([A-Za-z]+)", str(cell_range))
                                if not m: continue
                                c = column_index_from_string(m.group(1)) - 1
                            if 0 <= c < len(df_cols):
                                dropdowns[df_cols[c]] = options
                        except:
                            continue
    return dropdowns

def whatsapp_link(phone, msg):
    return f"https://wa.me/{phone}?text={urllib.parse.quote(msg)}"

# ==============================
# CHECK FOR MEMBER MODE
# ==============================
params = st.experimental_get_query_params()
form_id = params.get("form_id", [None])[0]
meta = load_meta()

# ==============================
# MEMBER FORM MODE
# ==============================
if form_id:
    if form_id not in meta:
        st.error("Invalid or expired form link.")
        st.stop()

    form_info = meta[form_id]
    form_name = form_info["form_name"]
    cols = form_info["columns"]
    dropdowns = form_info.get("dropdowns", {})
    st.header(f"📋 {form_name}")

    member_name = st.text_input("Your Name:")
    form_data = {}
    for col in cols:
        if col.lower() == "name": continue
        if col in dropdowns:
            form_data[col] = st.selectbox(col, dropdowns[col])
        else:
            form_data[col] = st.text_input(col)

    submit = st.button("✅ Submit Form")
    if submit:
        if not member_name:
            st.error("Please enter your name.")
        else:
            # Save submission permanently
            sub_path = os.path.join(DATA_DIR, f"submissions_{form_id}.xlsx")
            new_entry = {"Name": member_name, **form_data, "SubmittedAt": datetime.now().strftime("%Y-%m-%d %H:%M")}
            if os.path.exists(sub_path):
                df = pd.read_excel(sub_path)
                df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
            else:
                df = pd.DataFrame([new_entry])
            df.to_excel(sub_path, index=False)

            # Update member status
            status_path = os.path.join(DATA_DIR, f"members_{form_id}.xlsx")
            if os.path.exists(status_path):
                ms = pd.read_excel(status_path)
                mask = ms["Name"].astype(str).str.lower() == member_name.strip().lower()
                if mask.any():
                    ms.loc[mask, "Status"] = "✅ Filled"
                    ms.loc[mask, "LastSubmitted"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                else:
                    ms = pd.concat([ms, pd.DataFrame([{
                        "Name": member_name, "Whatsapp": "", "Status": "✅ Filled",
                        "LastSubmitted": datetime.now().strftime("%Y-%m-%d %H:%M")
                    }])], ignore_index=True)
            else:
                ms = pd.DataFrame([{
                    "Name": member_name, "Whatsapp": "", "Status": "✅ Filled",
                    "LastSubmitted": datetime.now().strftime("%Y-%m-%d %H:%M")
                }])
            ms.to_excel(status_path, index=False)
            st.success("🎉 Your response has been recorded. Thank you!")

else:
    # ==============================
    # ADMIN PANEL
    # ==============================
    st.header("🧑‍💼 Admin Panel")

    members_file = st.file_uploader("👥 Upload Members Excel (must have Name, Whatsapp)", type=["xlsx"])
    form_file = st.file_uploader("📄 Upload Form Excel Template", type=["xlsx"])

    if members_file and form_file:
        # Load members
        members = pd.read_excel(members_file)
        members.columns = [c.title().strip() for c in members.columns]
        if not {"Name", "Whatsapp"}.issubset(set(members.columns)):
            st.error("Members file must contain 'Name' and 'Whatsapp' columns.")
            st.stop()

        members["Whatsapp"] = members["Whatsapp"].apply(normalize_phone)
        members["Status"] = "❌ Pending"
        members["LastSubmitted"] = ""

        # Load form
        hdr = detect_header_row(form_file)
        form_df = pd.read_excel(form_file, header=hdr)
        form_df.columns = [str(c).strip().replace("_", " ").title() for c in form_df.columns]
        dropdowns = detect_dropdowns(form_file, list(form_df.columns))
        st.success(f"Detected {len(form_df.columns)} form fields ✅")
        st.write(form_df.columns.tolist())

        # Generate form ID & link
        form_id = str(uuid.uuid4())[:10]
        form_name = st.text_input("Form Name:", f"My Form {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        base_url = st.text_input("Your Streamlit App URL (e.g. https://myapp.streamlit.app):")

        if st.button("🚀 Generate WhatsApp Links + Save Form"):
            meta = load_meta()
            meta[form_id] = {
                "form_name": form_name,
                "columns": list(form_df.columns),
                "dropdowns": dropdowns,
                "created_at": datetime.now().isoformat()
            }
            save_meta(meta)

            # Save members status
            members_path = os.path.join(DATA_DIR, f"members_{form_id}.xlsx")
            members.to_excel(members_path, index=False)

            form_link = f"{base_url}/?form_id={form_id}"
            st.info(f"🔗 Form Link: {form_link}")

            # Generate WhatsApp links
            links = []
            for _, r in members.iterrows():
                name, phone = r["Name"], r["Whatsapp"]
                msg = f"Hello {name}! Please fill your form here: {form_link}"
                wa = whatsapp_link(phone, msg)
                links.append((name, phone, wa))

            wa_df = pd.DataFrame(links, columns=["Name", "Whatsapp", "WhatsApp Link"])
            st.dataframe(wa_df)

            buf = BytesIO()
            wa_df.to_excel(buf, index=False)
            st.download_button("⬇️ Download WhatsApp Links Excel", data=buf.getvalue(),
                               file_name=f"wa_links_{form_id}.xlsx")

            st.success("✅ Form saved permanently. Members can now receive their links!")

    # ==============================
    # VIEW EXISTING FORMS
    # ==============================
    st.subheader("📊 Existing Forms & Progress")
    meta = load_meta()
    if not meta:
        st.info("No forms created yet.")
    else:
        all_forms = []
        for fid, info in meta.items():
            sub_path = os.path.join(DATA_DIR, f"submissions_{fid}.xlsx")
            ms_path = os.path.join(DATA_DIR, f"members_{fid}.xlsx")
            filled = 0
            total = 0
            if os.path.exists(ms_path):
                ms = pd.read_excel(ms_path)
                total = len(ms)
                filled = (ms["Status"] == "✅ Filled").sum()
            all_forms.append({
                "Form ID": fid,
                "Name": info["form_name"],
                "Created": info["created_at"][:10],
                "Filled": filled,
                "Total": total
            })
        st.dataframe(pd.DataFrame(all_forms))
