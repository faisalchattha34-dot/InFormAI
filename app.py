import streamlit as st
import pandas as pd
import uuid
from datetime import datetime
from io import BytesIO

# ----------------------------
# SESSION STATE
# ----------------------------
if "forms" not in st.session_state:
    st.session_state.forms = {}  # {form_id: {...}}
if "responses" not in st.session_state:
    st.session_state.responses = []  # list of dicts
if "current_form_id" not in st.session_state:
    st.session_state.current_form_id = None
if "members" not in st.session_state:
    st.session_state.members = {}  # form_id -> list of emails
if "edit_response_idx" not in st.session_state:
    st.session_state.edit_response_idx = None

# ----------------------------
# PAGE NAVIGATION
# ----------------------------
st.set_page_config(page_title="Formlify SaaS", layout="wide")
mode = st.sidebar.radio("Mode", ["Admin", "User"])

# ----------------------------
# HELPER FUNCTIONS
# ----------------------------

def detect_columns(excel_file):
    df = pd.read_excel(excel_file)
    columns = []
    for c in df.columns:
        columns.append({"label": str(c), "type": "Text", "required": False, "options": ""})
    return columns

def get_user_form_link(form_id, user_email):
    # Unique link simulation
    return f"{st.get_url()}?mode=User&form_id={form_id}&user_email={user_email}"

# ----------------------------
# ADMIN MODE
# ----------------------------
if mode == "Admin":
    st.title("🧑‍💼 Admin Panel")

    # ------------------------
    # Upload Form + Members
    # ------------------------
    st.subheader("📄 Create New Form")
    form_name = st.text_input("Form Name", f"My Form {datetime.now().strftime('%Y-%m-%d')}")

    col1, col2 = st.columns(2)
    with col1:
        form_file = st.file_uploader("Upload Form Excel", type=["xlsx"])
    with col2:
        members_file = st.file_uploader("Upload Members Email List", type=["xlsx"])

    if st.button("🚀 Create Form & Detect Columns"):
        if not form_file or not members_file:
            st.error("Please upload both files.")
        else:
            # Detect form columns
            columns = detect_columns(form_file)
            # Load members
            df_members = pd.read_excel(members_file)
            emails = df_members["Email"].dropna().unique().tolist()
            form_id = str(uuid.uuid4())[:8]
            st.session_state.forms[form_id] = {
                "name": form_name,
                "columns": columns,
                "created_at": datetime.now(),
            }
            st.session_state.members[form_id] = emails
            st.session_state.current_form_id = form_id
            st.success(f"Form created! ID: {form_id}")

    # ------------------------
    # Column & Form Editing
    # ------------------------
    if st.session_state.current_form_id:
        st.subheader("✏️ Form Columns")
        columns = st.session_state.forms[st.session_state.current_form_id]["columns"]
        for i, col in enumerate(columns):
            c1, c2, c3, c4, c5 = st.columns([3,2,1,1,1])
            col["label"] = c1.text_input("Label", col["label"], key=f"label_{i}")
            col["type"] = c2.selectbox("Type", ["Text","Number","Dropdown"], index=["Text","Number","Dropdown"].index(col["type"]), key=f"type_{i}")
            if col["type"] == "Dropdown":
                col["options"] = c3.text_input("Options (comma separated)", col["options"], key=f"opt_{i}")
            col["required"] = c4.checkbox("Required", col["required"], key=f"req_{i}")
            if c5.button("🗑 Delete", key=f"del_{i}"):
                columns.pop(i)
                st.experimental_rerun()
        if st.button("➕ Add New Column"):
            columns.append({"label":"New Field","type":"Text","required":False,"options":""})
            st.experimental_rerun()

        st.markdown("### 👀 Form Preview")
        with st.form("preview_form"):
            for f in columns:
                label = f"{f['label']} {'*' if f['required'] else ''}"
                if f["type"]=="Text":
                    st.text_input(label)
                elif f["type"]=="Number":
                    st.number_input(label)
                elif f["type"]=="Dropdown":
                    opts = [o.strip() for o in f["options"].split(",") if o.strip()]
                    st.selectbox(label, opts if opts else ["Option 1"])
            st.form_submit_button("Submit (Preview)")

    # ------------------------
    # Responses Dashboard
    # ------------------------
    st.subheader("📊 Responses Dashboard")
    if st.session_state.responses:
        df_resp = pd.DataFrame(st.session_state.responses)
        st.dataframe(df_resp)
        # Edit/Delete simulation
        for idx, row in df_resp.iterrows():
            col1, col2 = st.columns([1,1])
            if col1.button(f"✏️ Edit {idx}"):
                st.session_state.edit_response_idx = idx
            if col2.button(f"🗑 Delete {idx}"):
                st.session_state.responses.pop(idx)
                st.experimental_rerun()
        if st.session_state.edit_response_idx is not None:
            st.markdown(f"### Editing Response #{st.session_state.edit_response_idx}")
            row = st.session_state.responses[st.session_state.edit_response_idx]
            for key in row:
                row[key] = st.text_input(key, str(row[key]))
            if st.button("💾 Save Changes"):
                st.session_state.responses[st.session_state.edit_response_idx] = row
                st.session_state.edit_response_idx = None
                st.experimental_rerun()

# ----------------------------
# USER MODE
# ----------------------------
elif mode == "User":
    params = st.experimental_get_query_params()
    form_id = params.get("form_id",[None])[0]
    user_email = params.get("user_email",[None])[0]

    if not form_id or form_id not in st.session_state.forms:
        st.warning("Invalid form ID.")
    else:
        form = st.session_state.forms[form_id]
        st.title(f"🧾 {form['name']}")
        st.subheader(f"Welcome: {user_email}")

        with st.form("user_fill_form"):
            user_data = {}
            for f in form["columns"]:
                label = f["label"] + (" *" if f["required"] else "")
                if f["type"]=="Text":
                    user_data[label] = st.text_input(label)
                elif f["type"]=="Number":
                    user_data[label] = st.number_input(label)
                elif f["type"]=="Dropdown":
                    opts = [o.strip() for o in f["options"].split(",") if o.strip()]
                    user_data[label] = st.selectbox(label, opts if opts else ["Option 1"])
            submitted = st.form_submit_button("✅ Submit Form")
            if submitted:
                st.session_state.responses.append({
                    "FormID": form_id,
                    "Email": user_email,
                    "SubmittedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    **user_data
                })
                st.success("Form submitted successfully!")
