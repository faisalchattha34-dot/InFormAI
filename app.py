import streamlit as st
import pandas as pd
import uuid
import os
from datetime import datetime

st.set_page_config(page_title="Auto Form System", layout="wide")

# ================================================================
# STORAGE FILES
# ================================================================
MEMBERS_FILE = "members.xlsx"
FORMS_META = "form_meta.json"
RESPONSES_FILE = "responses.xlsx"

# ================================================================
# UTILs
# ================================================================
def load_excel(f):
    if os.path.exists(f):
        return pd.read_excel(f)
    return pd.DataFrame()

def save_excel(df, f):
    df.to_excel(f, index=False)

# ================================================================
# LOAD DATA
# ================================================================
members_df = load_excel(MEMBERS_FILE)
responses_df = load_excel(RESPONSES_FILE)

def ensure_responses_columns():
    cols = ["FormID", "SubmittedAt"]
    for c in members_df.columns:
        if c not in cols:
            cols.append(c)
    for c in cols:
        if c not in responses_df.columns:
            responses_df[c] = ""
    save_excel(responses_df, RESPONSES_FILE)

ensure_responses_columns()

# ================================================================
# SIDEBAR MENU
# ================================================================
menu = st.sidebar.radio("📌 Menu", ["Upload Files", "Create Form", "Responses"])

# ================================================================
# UPLOAD FILES
# ================================================================
if menu == "Upload Files":
    st.header("📤 Upload Excel Files")

    c1, c2 = st.columns(2)

    with c1:
        uploaded_members = st.file_uploader("Upload Member Emails File", type=["xlsx"])
        if uploaded_members:
            df = pd.read_excel(uploaded_members)
            save_excel(df, MEMBERS_FILE)
            st.success("Members List Saved")

    with c2:
        uploaded_form = st.file_uploader("Upload Form Structure File", type=["xlsx"])
        if uploaded_form:
            df = pd.read_excel(uploaded_form)
            df.to_excel("form_structure.xlsx", index=False)
            st.success("Form Structure Saved")

    st.info("📌 Done — Now go to Create Form")

# ================================================================
# CREATE FORM & SEND LINK
# ================================================================
elif menu == "Create Form":
    st.header("🧾 Create Form")

    if not os.path.exists("form_structure.xlsx"):
        st.warning("Upload Form Structure First")
    else:
        form_df = pd.read_excel("form_structure.xlsx")
        st.subheader("Detected Form Fields")
        st.dataframe(form_df)

        form_id = uuid.uuid4().hex[:6]
        link = f"https://yourapp.streamlit.app/?form={form_id}"

        st.success(f"🔗 Form Link: {link}")

# ================================================================
# RESPONSES TABLE + EDIT + DELETE
# ================================================================
elif menu == "Responses":
    st.header("📊 Form Responses")

    if responses_df.empty:
        st.info("No Responses Found Yet")
        st.stop()

    search = st.text_input("🔍 Search Value")
    temp_df = responses_df

    if search:
        temp_df = temp_df[temp_df.apply(lambda row: row.astype(str).str.contains(search, case=False).any(), axis=1)]

   PAGE_SIZE = 10
    total = len(temp_df)
    pages = (total // PAGE_SIZE) + (1 if total % PAGE_SIZE else 0)

    if "page" not in st.session_state:
        st.session_state.page = 1

    col1, col2 = st.columns([1,1])
    if col1.button("⬅ Previous") and st.session_state.page > 1:
        st.session_state.page -= 1
    if col2.button("Next ➡") and st.session_state.page < pages:
        st.session_state.page += 1

    start = (st.session_state.page - 1) * PAGE_SIZE
    end = start + PAGE_SIZE
    page_df = temp_df.iloc[start:end]

    st.write(f"📄 Page {st.session_state.page}/{pages}")
    st.dataframe(page_df, use_container_width=True)

    st.subheader("🗑 Multi Delete")
    del_rows = st.multiselect("Select Rows to DELETE", list(page_df.index))
    if st.button("Delete Selected"):
        responses_df.drop(del_rows, inplace=True)
        responses_df.reset_index(drop=True, inplace=True)
        save_excel(responses_df, RESPONSES_FILE)
        st.success("Selected rows deleted")

    st.subheader("✏ Edit Response")

    if not page_df.empty:
        edit_idx = st.selectbox("Select Row Index to EDIT", list(page_df.index))
        row = responses_df.loc[edit_idx]
        new_vals = {}
        for col in responses_df.columns:
            if col not in ["FormID", "SubmittedAt"]:
                new_vals[col] = st.text_input(col, value=row[col])

        if st.button("💾 Save Edit"):
            for k, v in new_vals.items():
                responses_df.at[edit_idx, k] = v
            responses_df.at[edit_idx, "LastEdited"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            save_excel(responses_df, RESPONSES_FILE)
            st.success("Row Updated Successfully")

    st.subheader("⬇ Download Excel")
    st.download_button(
        "Download Responses Excel",
        data=open(RESPONSES_FILE, "rb").read(),
        file_name="responses.xlsx"
    )
