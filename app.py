import streamlit as st
import pandas as pd

st.set_page_config(page_title="Form Builder", layout="wide")

st.title("🧩 Excel → Form Builder")

# ----------------------------
# SESSION STATE
# ----------------------------
if "fields" not in st.session_state:
    st.session_state.fields = []

# ----------------------------
# EXCEL UPLOAD
# ----------------------------
uploaded = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    st.success("Excel loaded successfully!")

    if not st.session_state.fields:
        for col in df.columns:
            st.session_state.fields.append({
                "label": col,
                "type": "Text",
                "required": False,
                "options": ""
            })

# ----------------------------
# FIELD EDITOR
# ----------------------------
st.markdown("## ✏️ Form Fields Editor")

for i, field in enumerate(st.session_state.fields):
    with st.container():
        c1, c2, c3, c4, c5 = st.columns([3,2,2,1,1])

        field["label"] = c1.text_input(
            "Label", field["label"], key=f"label_{i}"
        )

        field["type"] = c2.selectbox(
            "Type",
            ["Text", "Number", "Dropdown"],
            index=["Text", "Number", "Dropdown"].index(field["type"]),
            key=f"type_{i}"
        )

        if field["type"] == "Dropdown":
            field["options"] = c3.text_input(
                "Options (comma separated)",
                field["options"],
                key=f"opt_{i}"
            )
        else:
            c3.markdown("—")

        field["required"] = c4.checkbox(
            "Required", field["required"], key=f"req_{i}"
        )

        # Drag-like ordering
        if c5.button("⬆", key=f"up_{i}") and i > 0:
            st.session_state.fields[i-1], st.session_state.fields[i] = \
                st.session_state.fields[i], st.session_state.fields[i-1]
            st.experimental_rerun()

        if c5.button("⬇", key=f"down_{i}") and i < len(st.session_state.fields)-1:
            st.session_state.fields[i+1], st.session_state.fields[i] = \
                st.session_state.fields[i], st.session_state.fields[i+1]
            st.experimental_rerun()

        st.divider()

# ----------------------------
# ADD NEW FIELD
# ----------------------------
st.button("➕ Add New Field", on_click=lambda: st.session_state.fields.append({
    "label": "New Field",
    "type": "Text",
    "required": False,
    "options": ""
}))

# ----------------------------
# LIVE PREVIEW
# ----------------------------
st.markdown("## 👀 Live Form Preview")

with st.form("preview_form"):
    for f in st.session_state.fields:
        label = f["label"] + (" *" if f["required"] else "")

        if f["type"] == "Text":
            st.text_input(label)
        elif f["type"] == "Number":
            st.number_input(label)
        elif f["type"] == "Dropdown":
            opts = [o.strip() for o in f["options"].split(",") if o.strip()]
            st.selectbox(label, opts if opts else ["Option 1"])

    st.form_submit_button("Submit (Preview)")
