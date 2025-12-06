import streamlit as st
import pandas as pd
import uuid
import os
import smtplib
from email.mime.text import MIMEText

st.set_page_config(page_title="📄 Editable Form", layout="wide")
st.title("📄 Editable Form + Edit/Delete + Search + Pagination")

DATA_FILE = "responses.xlsx"

# =====================================
# LOAD / SAVE
# =====================================
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_excel(DATA_FILE)
    return pd.DataFrame(columns=["id", "Name", "Email", "City"])

def save_data(df):
    df.to_excel(DATA_FILE, index=False)

df = load_data()


# =====================================
# EMAIL SENDING FUNCTION
# =====================================
def send_email(to, subject, body):
    try:
        sender_email = st.session_state.get("smtp_user")
        sender_pass = st.session_state.get("smtp_pass")

        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = to

        smtp = smtplib.SMTP(st.session_state.get("smtp_server"), st.session_state.get("smtp_port"))
        smtp.starttls()
        smtp.login(sender_email, sender_pass)
        smtp.sendmail(sender_email, to, msg.as_string())
        smtp.quit()
        return True
    except:
        return False


# =====================================
# SIDEBAR SETTINGS
# =====================================
st.sidebar.header("🔍 Search")
search_query = st.sidebar.text_input("Search by Name/Email/City")

st.sidebar.header("📨 Email Settings (optional)")
smtp_enable = st.sidebar.checkbox("Enable Email Send?")
if smtp_enable:
    st.session_state.smtp_server = st.sidebar.text_input("SMTP Server", value="smtp.gmail.com")
    st.session_state.smtp_port = st.sidebar.number_input("Port", value=587)
    st.session_state.smtp_user = st.sidebar.text_input("Email Address")
    st.session_state.smtp_pass = st.sidebar.text_input("Password", type="password")

# =====================================
# FILTER SEARCH
# =====================================
if search_query:
    df = df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)]

# =====================================
# FORM (ADD / EDIT)
# =====================================
st.subheader("📝 Fill Form")

edit_id = st.session_state.get("edit_id", None)

if edit_id:
    rec = df[df["id"] == edit_id].iloc[0]
    default_name = rec["Name"]
    default_email = rec["Email"]
    default_city = rec["City"]
else:
    default_name = default_email = default_city = ""

name = st.text_input("Full Name", default_name)
email = st.text_input("Email", default_email)
city = st.text_input("City", default_city)

if st.button("Save"):
    if edit_id:
        df.loc[df["id"] == edit_id, ["Name", "Email", "City"]] = [name, email, city]
        st.session_state.edit_id = None
        st.success("✔ Updated Successfully!")
        if smtp_enable:
            send_email(email, "✔ Record Updated", f"Your information updated successfully.\n\nName: {name}\nCity: {city}")
    else:
        new_row = {"id": str(uuid.uuid4()), "Name": name, "Email": email, "City": city}
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        st.success("✔ Added Successfully!")
        if smtp_enable:
            send_email(email, "✔ Record Added", f"Thank you for submitting.\n\nName: {name}\nCity: {city}")

    save_data(df)
    st.rerun()


# =====================================
# Download Excel
# =====================================
st.subheader("📁 Download")
st.download_button(
    label="⬇ Download Excel",
    data=open(DATA_FILE, "rb").read(),
    file_name="responses.xlsx",
    mime="application/vnd.ms-excel"
)


# =====================================
# PAGINATION
# =====================================
PAGE_SIZE = 5
total_records = len(df)
total_pages = max(1, (total_records + PAGE_SIZE - 1) // PAGE_SIZE)

if "page" not in st.session_state:
    st.session_state.page = 1

# Prev/Next Buttons
prev_col, mid, next_col = st.columns([1,3,1])
if prev_col.button("⬅ Previous") and st.session_state.page > 1:
    st.session_state.page -= 1
if next_col.button("Next ➡") and st.session_state.page < total_pages:
    st.session_state.page += 1

start = (st.session_state.page - 1) * PAGE_SIZE
end = start + PAGE_SIZE
df_show = df.iloc[start:end]


# =====================================
# DISPLAY DATA
# =====================================
st.subheader(f"📋 Records (Page {st.session_state.page}/{total_pages})")

if df_show.empty:
    st.info("No Records Found")
else:
    for index, row in df_show.iterrows():
        col1, col2, col3 = st.columns([6,1,1])
        col1.write(f"**{row['Name']}** | {row['Email']} | {row['City']}")

        # Edit Button
        if col2.button("✏️ Edit", key=f"edit{row['id']}"):
            st.session_state.edit_id = row["id"]
            st.rerun()

        # Delete Button
        if col3.button("🗑 Delete", key=f"delete{row['id']}"):
            df = df[df["id"] != row["id"]]
            save_data(df)
            st.rerun()
