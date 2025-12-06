import streamlit as st
import pandas as pd
import os
import json
import uuid
from datetime import datetime
from openpyxl import load_workbook
from io import BytesIO
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ----------------------------
# Setup
# ----------------------------
st.set_page_config(page_title="📄 Excel → Web Form + Auto Email", layout="wide")
st.title("📄 Excel → Web Form + Auto Email Sender + Dashboard")

# ----------------------------
# CSS Styling
# ----------------------------
st.markdown("""
<style>
:root { color-scheme: light dark; }
body { background-color: var(--background-color); font-family: 'Arial', sans-serif; }
h1,h2,h3,p,label,span,div { color: inherit !important; }
[data-baseweb="input"] input, [data-baseweb="select"] select { color: inherit !important; background-color: transparent !important; }
.stTextInput,.stSelectbox,.stTextArea,.stDataFrame { border-radius:8px; padding:10px; }
.stButton>button { background-color:#3498db;color:white;padding:10px 20px;border-radius:8px;border:none;font-size:16px;font-weight:500;transition:all 0.3s ease;}
.stButton>button:hover { background-color:#2980b9; transform: scale(1.03);}
.stDownloadButton>button { background-color:#2ecc71;color:white;padding:10px 20px;border-radius:8px;border:none;font-size:16px;font-weight:500;transition:all 0.3s ease;}
.stDownloadButton>button:hover { background-color:#27ae60; transform: scale(1.03);}
</style>
""", unsafe_allow_html=True)

# ----------------------------
# Paths & Helpers
# ----------------------------
DATA_DIR = "data_store"
os.makedirs(DATA_DIR, exist_ok=True)
META_PATH = os.path.join(DATA_DIR, "meta.json")
ALL_RESPONSES_PATH = os.path.join(DATA_DIR, "all_responses.xlsx")

def load_meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_meta(meta):
    with open(META_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

def detect_dropdowns(excel_file, df_columns):
    excel_file.seek(0)
    wb = load_workbook(excel_file, data_only=True)
    ws = wb.active
    dropdowns = {}
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            try:
                if dv.type=="list" and dv.formula1:
                    formula=str(dv.formula1).strip('"')
                    options=[x.strip() for x in formula.split(",")] if "," in formula else []
                    for cell_range in dv.cells:
                        cidx=cell_range.min_col-1
                        if 0<=cidx<len(df_columns):
                            dropdowns[df_columns[cidx]]=options
            except: continue
    return dropdowns

def send_email_to_members(sender_email,password,members,subject,message):
    sent_count=0
    results=[]
    for email in members:
        try:
            msg=MIMEMultipart()
            msg["From"]=sender_email
            msg["To"]=email
            msg["Subject"]=subject
            msg.attach(MIMEText(message,"plain"))
            with smtplib.SMTP("smtp.gmail.com",587) as server:
                server.starttls()
                server.login(sender_email,password)
                server.send_message(msg)
            sent_count+=1
            results.append({"Email":email,"Status":"✅ Sent"})
        except Exception as e:
            results.append({"Email":email,"Status":f"❌ Failed ({e})"})
    return sent_count, results

def load_responses():
    if os.path.exists(ALL_RESPONSES_PATH):
        return pd.read_excel(ALL_RESPONSES_PATH)
    return pd.DataFrame()

def save_responses(df):
    df.to_excel(ALL_RESPONSES_PATH,index=False)

# ----------------------------
# URL Params
# ----------------------------
params = st.experimental_get_query_params()
mode = params.get("mode",["admin"])[0]
form_id = params.get("form_id",[None])[0]
meta = load_meta()

# ----------------------------
# FORM VIEW
# ----------------------------
if mode=="form":
    if not form_id or "forms" not in meta or form_id not in meta["forms"]:
        st.warning("Invalid or missing form ID. Please select a form from below:")
        if "forms" in meta:
            for fid,f in meta["forms"].items():
                link=f"?mode=form&form_id={fid}"
                st.markdown(f"- [{f['form_name']}]({link})")
    else:
        info = meta["forms"][form_id]
        st.header(f"🧾 {info['form_name']}")
        if "session_id" not in st.session_state:
            st.session_state["session_id"]=str(uuid.uuid4())[:8]
        session_id = st.session_state["session_id"]
        dropdowns = info.get("dropdowns",{})
        columns = info["columns"]

        with st.form("user_form", clear_on_submit=False):
            values={}
            for col in columns:
                if col in dropdowns:
                    values[col]=st.selectbox(col,dropdowns[col], key=f"{col}_{session_id}")
                else:
                    values[col]=st.text_input(col,key=f"{col}_{session_id}")
            submitted=st.form_submit_button("✅ Submit Response")

        if submitted:
            row={
                "FormID":form_id,
                "FormName":info["form_name"],
                "UserSession":session_id,
                "SubmittedAt":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            row.update(values)
            try:
                responses=load_responses()
                for col in row.keys():
                    if col not in responses.columns:
                        responses[col]=None
                responses=pd.concat([responses,pd.DataFrame([row])],ignore_index=True)
                save_responses(responses)
                st.success("🎉 Response saved successfully!")
                st.balloons()
            except Exception as e:
                st.error(f"❌ Error saving data: {e}")

# ----------------------------
# ADMIN VIEW
# ----------------------------
else:
    st.header("🧑‍💼 Admin Panel")
    st.write("Upload two Excel files — Member List & Form Source.")
    col1,col2 = st.columns(2)
    with col1:
        member_file=st.file_uploader("📋 Upload Member List (must have 'Email' column)", type=["xlsx"])
    with col2:
        form_file=st.file_uploader("📄 Upload Form Source File", type=["xlsx"])

    if member_file and form_file:
        try:
            df_members=pd.read_excel(member_file)

            excel_data=pd.read_excel(form_file, header=None)
            header_row_index=None
            for i,row in excel_data.iterrows():
                if row.count()>=len(row)/2:
                    header_row_index=i
                    break
            df_form=pd.read_excel(form_file, header=header_row_index if header_row_index is not None else 0)

            # Clean header names
            cleaned_cols=[]
            seen=set()
            prev_name=None
            for c in df_form.columns:
                name=str(c).strip() if pd.notna(c) and str(c).strip() else prev_name
                if name:
                    name=name.replace("_"," ").title()
                    if name in seen:
                        i=2
                        while f"{name}_{i}" in seen: i+=1
                        name=f"{name}_{i}"
                    seen.add(name)
                    cleaned_cols.append(name)
                    prev_name=name
            df_form.columns=cleaned_cols

            # Editable preview
            st.subheader("👀 Edit Form Data (Live Preview)")
            if "original_columns" not in st.session_state:
                st.session_state.original_columns=list(df_form.columns)
            if "current_form_df" not in st.session_state:
                st.session_state.current_form_df=df_form.copy()

            edited_df=st.data_editor(
                st.session_state.current_form_df,
                use_container_width=True,
                num_rows="dynamic",
                key="form_editor",
            )
            st.session_state.current_form_df=edited_df.copy()

            st.write("### ✏️ Column Management")
            col_action=st.radio("Select Action", ["None","Rename Column","Delete Column","Add Column","Restore Deleted Column"], horizontal=True)

            if col_action=="Rename Column":
                col_to_rename=st.selectbox("Select column to rename", st.session_state.current_form_df.columns)
                new_name=st.text_input("Enter new column name:")
                if st.button("✅ Rename Now"):
                    st.session_state.current_form_df.rename(columns={col_to_rename:new_name}, inplace=True)
                    st.success(f"Column renamed from '{col_to_rename}' → '{new_name}'")

            elif col_action=="Delete Column":
                col_to_delete=st.selectbox("Select column to delete", st.session_state.current_form_df.columns)
                if st.button("🗑️ Delete Column"):
                    st.session_state.current_form_df.drop(columns=[col_to_delete], inplace=True)
                    st.success(f"Column '{col_to_delete}' deleted.")

            elif col_action=="Add Column":
                new_col_name=st.text_input("Enter new column name:")
                if st.button("➕ Add Column"):
                    if new_col_name in st.session_state.current_form_df.columns:
                        st.warning("Column already exists.")
                    else:
                        st.session_state.current_form_df[new_col_name]=""
                        st.success(f"Column '{new_col_name}' added.")

            elif col_action=="Restore Deleted Column":
                deleted_cols=[c for c in st.session_state.original_columns if c not in st.session_state.current_form_df.columns]
                if deleted_cols:
                    col_to_restore=st.selectbox("Select deleted column to restore", deleted_cols)
                    if st.button("♻️ Restore Column"):
                        idx=st.session_state.original_columns.index(col_to_restore)
                        df=st.session_state.current_form_df
                        df.insert(loc=idx, column=col_to_restore, value="")
                        st.session_state.current_form_df=df
                        st.success(f"Column '{col_to_restore}' restored successfully.")
                else:
                    st.info("No deleted columns found to restore.")

            save_changes=st.button("💾 Save Changes to Original Excel File")
            if save_changes:
                try:
                    with BytesIO() as buffer:
                        st.session_state.current_form_df.to_excel(buffer,index=False)
                        buffer.seek(0)
                        with open(form_file.name,"wb") as f:
                            f.write(buffer.read())
                    st.success("✅ All changes saved back to the uploaded Excel file successfully!")
                except Exception as e:
                    st.error(f"❌ Failed to save: {e}")

            # Continue workflow
            if "Email" not in df_members.columns:
                st.error("❌ Member file must contain an 'Email' column.")
            else:
                dropdowns=detect_dropdowns(form_file, list(st.session_state.current_form_df.columns))
                st.success(f"✅ Form fields detected: {len(st.session_state.current_form_df.columns)}")
                st.write(st.session_state.current_form_df.columns.tolist())
                if dropdowns:
                    st.info("Detected dropdowns:")
                    st.table(pd.DataFrame([{"Field":k,"Options":", ".join(v)} for k,v in dropdowns.items()]))

                form_name=st.text_input("Form Name:", value=f"My Form {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                base_url=st.text_input("Your Streamlit App Public URL (example: https://yourapp.streamlit.app)")
                sender_email=st.text_input("Your Gmail Address:")
                password=st.text_input("Your Gmail App Password:", type="password")

                if st.button("🚀 Create Form & Send Emails"):
                    if not base_url:
                        st.error("Please enter your app URL.")
                    elif not sender_email or not password:
                        st.error("Please enter Gmail and App Password.")
                    else:
                        form_id_new=str(uuid.uuid4())[:10]
                        forms=meta.get("forms",{})
                        forms[form_id_new]={
                            "form_name":form_name,
                            "columns":list(st.session_state.current_form_df.columns),
                            "dropdowns":dropdowns,
                            "created_at":datetime.now().isoformat(),
                        }
                        meta["forms"]=forms
                        save_meta(meta)
                        link=f"{base_url.rstrip('/')}/?mode=form&form_id={form_id_new}"
                        st.success(f"✅ Form created successfully!\n{link}")
                        st.info("📧 Sending form link to all members...")
                        emails=df_members["Email"].dropna().unique().tolist()
                        subject=f"Form Invitation: {form_name}"
                        message=f"Hello,\n\nPlease fill out the form below:\n{link}\n\nThank you!"
                        sent_count,send_results=send_email_to_members(sender_email,password,emails,subject,message)
                        st.success(f"🎉 Emails sent: {sent_count}/{len(emails)}")
                        st.subheader("📧 Email Send Status")
                        st.table(pd.DataFrame(send_results))

        except Exception as e:
            st.error(f"❌ Error processing files: {e}")

    # ----------------------------
    # Responses Dashboard
    # ----------------------------
    st.markdown("---")
    st.subheader("📊 Responses Dashboard")
    responses=load_responses()
    if responses.empty:
        st.info("No responses submitted yet.")
    else:
        form_filter=st.selectbox("Select Form to View Responses:", ["All"]+[f["form_name"] for f in meta.get("forms",{}).values()])
        if form_filter!="All":
            form_id_list=[fid for fid,f in meta["forms"].items() if f["form_name"]==form_filter]
            responses_display=responses[responses["FormID"]==form_id_list[0]] if form_id_list else pd.DataFrame()
        else:
            responses_display=responses.copy()

        if not responses_display.empty:
            # Editable and deletable table
            st.write("### ✏️ Edit Responses")
            edited_responses = st.data_editor(
                responses_display,
                use_container_width=True,
                num_rows="dynamic",
                key="responses_editor",
            )
            # Save edits
            if st.button("💾 Save Edited Responses"):
                for idx, row in edited_responses.iterrows():
                    responses.loc[row.name] = row
                save_responses(responses)
                st.success("✅ Responses updated successfully!")

            # Delete single response
            st.markdown("### 🗑️ Delete a Response")
            idx_to_delete = st.number_input(
                "Select Row Number to Delete:",
                min_value=0,
                max_value=len(responses_display)-1,
                step=1
            )
            if st.button("❌ Delete Selected Response"):
                row_index = responses_display.index[idx_to_delete]
                responses = responses.drop(row_index).reset_index(drop=True)
                save_responses(responses)
                st.success("🗑️ Response deleted successfully!")
                st.experimental_rerun()

            # Hide metadata for display
            hidden_cols=["FormID","FormName","UserSession","SubmittedAt"]
            display_df=responses_display.drop(columns=[c for c in hidden_cols if c in responses_display.columns])
            st.dataframe(display_df,use_container_width=True)

            # Download responses
            to_download=BytesIO()
            display_df.to_excel(to_download,index=False)
            to_download.seek(0)
            st.download_button(
                label="📥 Download Responses",
                data=to_download,
                file_name="responses.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
