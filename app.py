import streamlit as st
import pandas as pd
import os
import uuid
from datetime import datetime

st.set_page_config(page_title="📄 Editable Responses App", layout="centered")
st.title("📄 Editable Form + Responses")

DATA_FILE = "responses.csv"

# ---------------------  Load Old Responses  ---------------------
if os.path.exists(DATA_FILE):
    df = pd.read_csv(DATA_FILE)
else:
    df = pd.DataFrame(columns=["id", "name", "email", "message", "timestamp"])


# ---------------------  Sidebar Menu  ---------------------
st.sidebar.header("📌 Menu")
menu = st.sidebar.radio("Select Option", ["Fill Form", "View Responses", "Edit Response"])


# ---------------------  Fill Form  ---------------------
if menu == "Fill Form":
    st.subheader("📝 Fill New Response")

    name = st.text_input("Name")
    email = st.text_input("Email")
    message = st.text_area("Message")

    if st.button("Submit"):
        new_row = {
            "id": str(uuid.uuid4()),
            "name": name,
            "email": email,
            "message": message,
            "timestamp": datetime.now()
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_csv(DATA_FILE, index=False)
        st.success("Response Saved Successfully!")


# ---------------------  View Responses  ---------------------
elif menu == "View Responses":
    st.subheader("📂 All Responses")
    if len(df) == 0:
        st.warning("No responses found!")
    else:
        st.dataframe(df)


# ---------------------  Edit Response  ---------------------
elif menu == "Edit Response":
    st.subheader("✏️ Edit an Existing Response")

    if len(df) == 0:
        st.warning("No responses available for edit!")
    else:
        selected_id = st.selectbox(
            "Select Response to Edit",
            options=df["id"],
            format_func=lambda x: f"{x} → {df[df['id']==x]['name'].values[0]}"
        )

        row = df[df["id"] == selected_id].iloc[0]

        new_name = st.text_input("Name", row["name"])
        new_email = st.text_input("Email", row["email"])
        new_message = st.text_area("Message", row["message"])

        if st.button("Update"):
            df.loc[df["id"] == selected_id, "name"] = new_name
            df.loc[df["id"] == selected_id, "email"] = new_email
            df.loc[df["id"] == selected_id, "message"] = new_message
            df.to_csv(DATA_FILE, index=False)
            st.success("Response Updated Successfully!")
