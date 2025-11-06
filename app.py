import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import os
import re
import urllib.parse
from datetime import datetime

st.set_page_config(page_title="📱 Excel Form + WhatsApp Automation", page_icon="📄", layout="centered")
st.title("📱 Excel Form + WhatsApp Link & Tracking System (Persistent Version)")

# --- Storage file (for persistent saving) ---
DATA_FILE = "form_submissions_data.xlsx"

# --- Upload Section ---
st.header("Step 1️⃣ : Upload Files")

members_file = st.file_uploader("👥 Upload Members List (must contain columns: Name, Whatsapp)", type=["xlsx"])
form_file = st.file_uploader("📄 Upload Form Layout Excel (columns define your form)", type=["xlsx"])

if members_file and form_file:
    try:
        # --- Read Members File ---
        members_df = pd.read_excel(members_file)
        if not {"Name", "Whatsapp"}.issubset(set(members_df.columns)):
            st.error("❌ Members file must have 'Name' and 'Whatsapp' columns.")
            st.stop()

        # --- Detect header row & read form file ---
        df_raw = pd.read_excel(form_file, header=None)
        header_row_index = None
        for i in range(len(df_raw)):
            row = df_raw.iloc[i]
            if row.notna().sum() > 2:
                header_row_index = i
                break
        df = pd.read_excel(form_file, header=header_row_index)
        df.columns = [str(c).strip().replace("_", " ").title() for c in df.columns if pd.notna(c)]

        st.success("✅ Form columns detected successfully!")
        st.write("**Form Columns:**", list(df.columns))

        # --- Detect dropdowns ---
        dropdown_dict = {}
        form_file.seek(0)
        wb = load_workbook(form_file, data_only=True)
        ws = wb.active
        if ws.data_validations:
            for dv in ws.data_validations.dataValidation:
                if dv.type == "list" and dv.formula1:
                    formula = str(dv.formula1).strip('"')
                    if "," in formula:
                        values = [v.strip() for v in formula.split(",")]
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
                                if 0 <= col_index < len(df.columns):
                                    dropdown_dict[df.columns[col_index]] = values
                            except Exception:
                                continue

        # --- Generate Form Link Automatically ---
        form_id = "F" + datetime.now().strftime("%Y%m%d_%H%M%S")
        base_link = f"https://your-app-name.streamlit.app/?form_id={form_id}"

        st.info(f"🔗 Generated Form Link: {base_link}")

        # --- Generate WhatsApp Message Links ---
        whatsapp_links = []
        for _, row in members_df.iterrows():
            name = str(row["Name"])
            phone = str(row["Whatsapp"]).replace("+", "").replace("-", "").replace(" ", "")
            msg = f"Hello {name}! Please fill your form here: {base_link}"
            encoded_msg = urllib.parse.quote(msg)
            wa_link = f"https://wa.me/{phone}?text={encoded_msg}"
            whatsapp_links.append(wa_link)

        members_df["Form Link"] = base_link
        members_df["WhatsApp Link"] = whatsapp_links
        members_df["Status"] = "❌ Pending"

        st.subheader("📲 Send Form Link to Members")
        st.dataframe(members_df[["Name", "Whatsapp", "Form Link", "WhatsApp Link", "Status"]])

        st.info("💡 Click the 'WhatsApp Link' to open message directly in WhatsApp and send it manually.")

        # --- Load Existing Saved Data if available ---
        if os.path.exists(DATA_FILE):
            saved_df = pd.read_excel(DATA_FILE)
        else:
            saved_df = pd.DataFrame(columns=df.columns)

        # --- Form Section ---
        st.subheader("🧾 Fill the Form")
        data = {}
        for col in df.columns:
            if col in dropdown_dict:
                data[col] = st.selectbox(f"{col}", dropdown_dict[col])
            else:
                data[col] = st.text_input(f"{col}")

        submit_btn = st.button("✅ Submit Form")
        if submit_btn:
            new_row = pd.DataFrame([data])
            saved_df = pd.concat([saved_df, new_row], ignore_index=True)
            saved_df.to_excel(DATA_FILE, index=False)
            st.success("🎉 Form submitted and permanently saved!")

        # --- Load Updated Data ---
        if os.path.exists(DATA_FILE):
            submissions_df = pd.read_excel(DATA_FILE)
        else:
            submissions_df = pd.DataFrame(columns=df.columns)

        st.subheader("📋 Submitted Data (Saved)")
        st.dataframe(submissions_df)

        # --- Track Progress ---
        total_members = len(members_df)
        filled = len(submissions_df)
        pending = total_members - filled
        progress = filled / total_members if total_members > 0 else 0

        st.progress(progress)
        st.write(f"✅ Filled: {filled} | ⏳ Pending: {pending} | Total: {total_members}")

        # --- Update Status for filled members ---
        for i, name in enumerate(members_df["Name"]):
            if name in submissions_df.get("Name", []):
                members_df.at[i, "Status"] = "✅ Filled"

        pending_members = members_df[members_df["Status"] == "❌ Pending"]["Name"].tolist()
        if pending_members:
            st.warning(f"❌ Pending Forms: {', '.join(pending_members)}")
        else:
            st.success("🎉 All members have submitted their forms!")

        # --- Download Combined Excel ---
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            members_df.to_excel(writer, index=False, sheet_name="Members Status")
            submissions_df.to_excel(writer, index=False, sheet_name="Form Submissions")
        buf.seek(0)

        st.download_button(
            label="⬇️ Download Excel (Members + Form Data)",
            data=buf,
            file_name="form_tracking_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
