import streamlit as st
import pandas as pd
import uuid
import os

st.set_page_config(page_title="📄 Editable Form", layout="centered")
st.title("📄 Editable Form with Edit/Delete")

DATA_FILE = "responses.xlsx"

# ---------------------------------
# Load or create excel file
# ---------------------------------
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_excel(DATA_FILE)
    else:
        return pd.DataFrame(columns=["id", "Name", "Email", "City"])

def save_data(df):
    df.to_excel(DATA_FILE, index=False)

# ---------------------------------
# INITIAL LOAD
# ---------------------------------
df = load_data()

# ---------------------------------
# ADD / EDIT FORM
# ---------------------------------
st.subheader("📝 Fill Form")

edit_id = st.session_state.get("edit_id", None)

if edit_id:
    rec = df[df["id"] == edit_id].iloc[0]
    default_name = rec["Name"]
    default_email = rec["Email"]
    default_city = rec["City"]
else:
    default_name = ""
    default_email = ""
    default_city = ""

name = st.text_input("Full Name", default_name)
email = st.text_input("Email", default_email)
city = st.text_input("City", default_city)

if st.button("Save"):
    if edit_id:
        df.loc[df["id"] == edit_id, ["Name", "Email", "City"]] = [name, email, city]
        st.session_state.edit_id = None
        st.success("✔ Record Updated Successfully!")
    else:
        new_row = {"id": str(uuid.uuid4()), "Name": name, "Email": email, "City": city}
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        st.success("✔ Record Added Successfully!")

    save_data(df)
    st.rerun()

# ---------------------------------
# SHOW DATA TABLE
# ---------------------------------
st.subheader("📋 All Records")

if df.empty:
    st.info("No Records Found")
else:
    for index, row in df.iterrows():
        col1, col2, col3 = st.columns([4,1,1])
        col1.write(f"**{row['Name']}** | {row['Email']} | {row['City']}")

        # EDIT BUTTON
        if col2.button("✏️ Edit", key=f"edit{row['id']}"):
            st.session_state.edit_id = row["id"]
            st.rerun()

        # DELETE BUTTON
        if col3.button("🗑 Delete", key=f"delete{row['id']}"):
            df = df[df["id"] != row["id"]]
            save_data(df)
            st.rerun()
