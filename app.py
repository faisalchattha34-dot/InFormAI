import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime

# -------------------------------------------------
# SETUP
# -------------------------------------------------
st.set_page_config(page_title="📋 Excel → Form System", layout="wide")
st.title("📋 Excel → Auto Web Form + Dashboard")

DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

META_FILE = os.path.join(DATA_DIR, "meta.json")
RESP_FILE = os.path.join(DATA_DIR, "all_responses.xlsx")

# -------------------------------------------------
# META HANDLING (FORMS CONFIG)
# -------------------------------------------------
def load_meta():
    if os.path.exists(META_FILE):
        try:
            return json.load(open(META_FILE, "r"))
        except:
            return {"forms": {}}
    return {"forms": {}}

def save_meta(meta):
    json.dump(meta, open(META_FILE, "w"), indent=4)

meta = load_meta()

# -------------------------------------------------
# RESPONSES EXCEL HANDLER
# -------------------------------------------------
def load_responses():
    if os.path.exists(RESP_FILE):
        return pd.read_excel(RESP_FILE)
    return pd.DataFrame()

def save_responses(df):
    df.to_excel(RESP_FILE, index=False)

# -------------------------------------------------
# UPLOAD NEW FORM
# -------------------------------------------------
st.sidebar.header("➕ Upload Excel Form")

uploaded = st.sidebar.file_uploader("Upload Excel File:")
if uploaded:
    df = pd.read_excel(uploaded)

    form_id = str(len(meta["forms"]) + 1)
    form_name = uploaded.name

    dropdowns = {}
    for col in df.columns:
        if df[col].dtype == "object":
            vals = df[col].dropna().unique().tolist()
            if 1 < len(vals) <= 20:
                dropdowns[col] = vals

    meta["forms"][form_id] = {
        "form_name": form_name,
        "columns": df.columns.tolist(),
        "dropdowns": dropdowns
    }
    save_meta(meta)
    st.sidebar.success("✅ Form Added Successfully!")

# -------------------------------------------------
# USER PANEL (FILL FORM)
# -------------------------------------------------
st.header("📝 Fill a Form")

if not meta["forms"]:
    st.info("No forms uploaded yet.")
else:
    selected_form = st.selectbox(
        "Select Form to Fill",
        [f["form_name"] for f in meta["forms"].values()]
    )

    form_id = [fid for fid, f in meta["forms"].items() if f["form_name"] == selected_form][0]
    form_info = meta["forms"][form_id]

    with st.form("fill_form"):
        values = {}
        for col in form_info["columns"]:
            if col in form_info["dropdowns"]:
                values[col] = st.selectbox(col, form_info["dropdowns"][col])
            else:
                values[col] = st.text_input(col)

        submit_form = st.form_submit_button("📤 Submit Form")

    if submit_form:
        df = load_responses()
        values["FormID"] = form_id
        values["FormName"] = selected_form
        values["SubmittedAt"] = datetime.now()
        df = pd.concat([df, pd.DataFrame([values])], ignore_index=True)
        save_responses(df)
        st.success("🎉 Form submitted successfully!")

# -------------------------------------------------
# ADMIN DASHBOARD
# -------------------------------------------------
st.markdown("---")
st.header("📊 Form Responses Dashboard")

responses = load_responses()
if responses.empty:
    st.info("No submissions yet.")
else:
    form_filter = st.selectbox(
        "Filter Responses by Form",
        ["All"] + [f["form_name"] for f in meta["forms"].values()]
    )

    if form_filter != "All":
        form_id_local = [fid for fid, f in meta["forms"].items() if f["form_name"] == form_filter][0]
        local_df = responses[responses["FormID"] == form_id_local]
    else:
        local_df = responses.copy()

    st.dataframe(local_df, use_container_width=True)

    # ------------ EDIT RESPONSE -------------------
    st.markdown("### ✏️ Edit a Response")
    if not local_df.empty:
        idx = st.number_input(
            "Select row number:",
            min_value=0,
            max_value=len(local_df) - 1,
            step=1
        )

        edit_button = st.button("✍️ Edit Selected Response")
        if edit_button:
            row = local_df.iloc[idx]
            form = meta["forms"].get(row["FormID"], {})

            st.subheader(f"Editing Response → {form['form_name']} ")

            with st.form("edit_response"):
                updated = {}
                for col in form["columns"]:
                    if col in form['dropdowns']:
                        opts = form['dropdowns'][col]
                        updated[col] = st.selectbox(
                            col, opts,
                            index=opts.index(row[col]) if row[col] in opts else 0
                        )
                    else:
                        updated[col] = st.text_input(col, value=row[col])

                save_edit = st.form_submit_button("💾 Save Changes")

            if save_edit:
                rid = row.name
                for col, val in updated.items():
                    responses.at[rid, col] = val
                save_responses(responses)
                st.success("✅ Response Updated!")
                st.experimental_rerun()

    # ------------ DELETE RESPONSE -------------------
    st.markdown("### 🗑️ Delete a Response")
    if not local_df.empty:
        del_idx = st.number_input(
            "Select row number to delete:",
            min_value=0,
            max_value=len(local_df) - 1,
            step=1
        )
        if st.button("❌ Delete Selected Response"):
            real_index = local_df.index[del_idx]
            responses = responses.drop(real_index).reset_index(drop=True)
            save_responses(responses)
            st.success("🗑️ Deleted Successfully!")
            st.experimental_rerun()

