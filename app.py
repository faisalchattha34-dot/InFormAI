import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re
import urllib.parse
import streamlit.components.v1 as components
from datetime import datetime

st.set_page_config(page_title="Excel Smart Form + WhatsApp (Manual Send)", page_icon="📄", layout="centered")
st.title("📄 Excel Smart Form + WhatsApp (Manual Send)")

# ---------------------------
# Helper functions
# ---------------------------
def detect_header_row(excel_file):
    df_raw = pd.read_excel(excel_file, header=None)
    header_row_index = None
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        if row.notna().sum() > 2:
            header_row_index = i
            break
    return header_row_index

def detect_dropdowns(form_file, df_columns):
    form_file.seek(0)
    wb = load_workbook(form_file, data_only=True)
    ws = wb.active
    dropdown_dict = {}
    if ws.data_validations is not None:
        for dv in ws.data_validations.dataValidation:
            try:
                if dv.type == "list" and dv.formula1:
                    formula = str(dv.formula1).strip('"')
                    # two types: inline comma-separated, or range reference
                    if "," in formula:
                        values = [v.strip() for v in formula.split(",")]
                    else:
                        # try to read range like Sheet!$A$1:$A$5 or $A$1:$A$5
                        values = []
                        try:
                            rng = formula.split("!")[-1].replace("$", "")
                            cells = rng.split(":")
                            if len(cells) == 2:
                                start = cells[0]
                                end = cells[1]
                                col_letters = re.match(r"([A-Za-z]+)", start).group(1)
                                start_row = int(re.match(r"[A-Za-z]+([0-9]+)", start).group(1))
                                end_row = int(re.match(r"[A-Za-z]+([0-9]+)", end).group(1))
                                col_index = column_index_from_string(col_letters)
                                for r in range(start_row, end_row + 1):
                                    cell_val = ws.cell(row=r, column=col_index).value
                                    if cell_val is not None:
                                        values.append(str(cell_val))
                        except Exception:
                            values = []
                    # find all affected columns (dv.cells)
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
                            if 0 <= col_index < len(df_columns):
                                dropdown_dict[df_columns[col_index]] = values
                        except Exception:
                            continue
            except Exception:
                continue
    return dropdown_dict

def make_whatsapp_url(number, message):
    # number should be digits, like 92300...
    encoded = urllib.parse.quote(message)
    return f"https://wa.me/{number}?text={encoded}"

def normalize_phone(ph):
    if pd.isna(ph):
        return ""
    s = str(ph).strip()
    s = s.replace("+", "").replace(" ", "").replace("-", "")
    # If number starts with 0, optionally replace leading 0 with country code?
    # We assume members provide full international number like 92300...
    return s

# ---------------------------
# Upload Members file
# ---------------------------
st.header("1) Upload Members List (Excel)")
st.write("Members Excel must have columns: `Name` and `Whatsapp` (e.g. 923001234567).")
members_file = st.file_uploader("Upload Members Excel", type=["xlsx"], key="members")

members_df = None
if members_file:
    try:
        members_df = pd.read_excel(members_file)
        # normalize columns
        members_df.columns = [str(c).strip().title() for c in members_df.columns]
        if "Name" not in members_df.columns or "Whatsapp" not in members_df.columns:
            st.error("Members file must contain columns: 'Name' and 'Whatsapp' (case-insensitive). Please fix and re-upload.")
            members_df = None
        else:
            # normalize phone column
            members_df["Whatsapp"] = members_df["Whatsapp"].apply(normalize_phone)
            # initial status
            if "Status" not in members_df.columns:
                members_df["Status"] = "❌ Pending"
            st.success("✅ Members loaded")
            st.dataframe(members_df[["Name", "Whatsapp", "Status"]].reset_index(drop=True))
    except Exception as e:
        st.error(f"Error reading members file: {e}")
        members_df = None

# ---------------------------
# Upload Form template
# ---------------------------
st.header("2) Upload Form Template (Excel)")
st.write("Upload the Excel file that contains form columns and any dropdowns you used in the template.")
form_file = st.file_uploader("Upload Form Excel", type=["xlsx"], key="form")

form_df = None
dropdowns = {}
if form_file:
    try:
        header_idx = detect_header_row(form_file)
        form_df = pd.read_excel(form_file, header=header_idx)
        # normalize column names
        form_df.columns = [str(c).strip().replace("_", " ").title() for c in form_df.columns if pd.notna(c)]
        st.success("✅ Form columns detected")
        st.write("Detected Columns:", list(form_df.columns))
        dropdowns = detect_dropdowns(form_file, list(form_df.columns))
        if dropdowns:
            st.info("Detected dropdowns:")
            st.table(pd.DataFrame([{"Column": c, "Options": ", ".join(v)} for c, v in dropdowns.items()]))
    except Exception as e:
        st.error(f"Error processing form file: {e}")
        form_df = None

# ---------------------------
# App link input (same link for everyone)
# ---------------------------
st.header("3) App Link (Same for all members)")
app_link = st.text_input("Enter the public Streamlit app link that members will open (same for everyone):",
                         placeholder="https://your-streamlit-app.streamlit.app")
st.caption("This is the link that will be included in WhatsApp messages. Make sure it's reachable by members.")

# ---------------------------
# Prepare session state containers
# ---------------------------
if "submissions" not in st.session_state:
    # store list of dicts for submissions
    st.session_state.submissions = []

if "members_status" not in st.session_state:
    # DataFrame with Name, Whatsapp, Status, LastSubmitted
    if members_df is not None:
        st.session_state.members_status = members_df[["Name", "Whatsapp", "Status"]].copy()
        st.session_state.members_status["LastSubmitted"] = ""
    else:
        st.session_state.members_status = pd.DataFrame(columns=["Name", "Whatsapp", "Status", "LastSubmitted"])

# If new members uploaded, update session state's members_status
if members_df is not None:
    # Rebuild members_status in session to reflect upload (but keep existing statuses if names match)
    existing = st.session_state.get("members_status", pd.DataFrame(columns=["Name", "Whatsapp", "Status", "LastSubmitted"]))
    new_status = []
    for _, r in members_df.iterrows():
        name = r["Name"]
        phone = r["Whatsapp"]
        # try to find in existing
        match = existing[existing["Name"] == name]
        if not match.empty:
            status = match.iloc[0]["Status"]
            last = match.iloc[0].get("LastSubmitted", "")
        else:
            status = "❌ Pending"
            last = ""
        new_status.append({"Name": name, "Whatsapp": phone, "Status": status, "LastSubmitted": last})
    st.session_state.members_status = pd.DataFrame(new_status)

# ---------------------------
# Show WhatsApp send links & send controls
# ---------------------------
if members_df is not None and app_link:
    st.header("4) WhatsApp Send Links (Manual send)")
    st.write("Click a member's 'Send' link to open WhatsApp Web/App with message ready. Or use 'Open All' (may be blocked by popup blockers).")
    # Build the message and links
    message_template = st.text_area("Message template (use {name} to insert member name). Example:",
                                    value="Hello {name}! Please fill your form here: {link}",
                                    height=80)
    if "{link}" not in message_template:
        st.error("Please include {link} in the message template so the app link is included.")
    else:
        rows = []
        wa_urls = []
        for _, r in st.session_state.members_status.iterrows():
            name = r["Name"]
            phone = r["Whatsapp"]
            msg = message_template.replace("{name}", name).replace("{link}", app_link)
            wa_url = make_whatsapp_url(phone, msg)
            wa_urls.append(wa_url)
            rows.append({
                "Name": name,
                "Whatsapp": phone,
                "Send Link": f"[Send to {name}]({wa_url})",
                "Status": r["Status"]
            })
        # Show as markdown table (clickable links)
        st.markdown(pd.DataFrame(rows).to_markdown(index=False), unsafe_allow_html=True)

        # "Open All" button using components.html (NOTE: browsers may block popups)
        if st.button("🚀 Open All WhatsApp Links (may be blocked by popup blocker)"):
            # generate JS to open each link in a new tab
            js = "<script>\n"
            js += "const links = " + str(wa_urls).replace("'", '"') + ";\n"
            js += "for(let i=0;i<links.length;i++){\n"
            js += "  window.open(links[i], '_blank');\n"
            js += "}\n"
            js += "</script>"
            components.html(js, height=10)

        st.info("Tip: If you plan to send messages one-by-one, click the individual 'Send to X' links. For bulk, try 'Open All' (may need to allow popups).")

# ---------------------------
# Form Filling Area (same link used by everyone)
# ---------------------------
if form_df is not None:
    st.header("5) Fill the Form (Members use this same form link)")
    st.write("Members must select their Name (so the system can mark them as 'Filled'). If you did not upload members file, they'll need to type their name.")

    # If we have members list, present a selectbox for 'Your name' to ensure matching
    if not st.session_state.members_status.empty:
        member_names = st.session_state.members_status["Name"].tolist()
        selected_name = st.selectbox("Your Name (select from list):", options=["-- I'm not listed --"] + member_names, key="selected_name")
        if selected_name == "-- I'm not listed --":
            # fallback to text input
            input_name = st.text_input("Please type your full name (as used by admin):", key="typed_name")
            submit_name = input_name.strip()
        else:
            submit_name = selected_name
    else:
        submit_name = st.text_input("Your Name:", key="typed_name_no_members").strip()

    # Render form fields
    st.write("Please fill the fields below and press Submit.")
    form_values = {}
    for col in form_df.columns:
        # avoid duplicate 'Name' colliding with selected_name field; if 'Name' is part of form, skip it because we capture above
        if col.lower() == "name":
            # skip to avoid confusion
            continue
        if col in dropdowns:
            form_values[col] = st.selectbox(col, dropdowns[col], key=f"f_{col}")
        else:
            form_values[col] = st.text_input(col, key=f"f_{col}")

    # Submit action
    if st.button("✅ Submit Form"):
        if not submit_name:
            st.error("Please provide your name so we can record who submitted the form.")
        else:
            # Save submission
            submission = {"_SubmittedBy": submit_name, "_SubmittedAt": datetime.utcnow().isoformat()}
            submission.update(form_values)
            st.session_state.submissions.append(submission)
            st.success("Thank you — your response has been recorded.")

            # mark member as filled if found in members_status
            ms = st.session_state.members_status
            match = ms[ms["Name"].str.strip().str.lower() == submit_name.strip().lower()]
            if not match.empty:
                idx = match.index[0]
                st.session_state.members_status.at[idx, "Status"] = "✅ Filled"
                st.session_state.members_status.at[idx, "LastSubmitted"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
            else:
                # if not in members list, append as a new row (optional)
                new_row = {"Name": submit_name, "Whatsapp": "", "Status": "✅ Filled", "LastSubmitted": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")}
                st.session_state.members_status = pd.concat([st.session_state.members_status, pd.DataFrame([new_row])], ignore_index=True)

# ---------------------------
# Show Submissions, Progress, Downloads
# ---------------------------
st.header("6) Submissions & Progress")

submissions_df = pd.DataFrame(st.session_state.submissions)
if not submissions_df.empty:
    st.subheader("All Submissions (latest first)")
    st.dataframe(submissions_df.iloc[::-1].reset_index(drop=True))
else:
    st.write("No submissions yet.")

# Show members status table
if not st.session_state.members_status.empty:
    st.subheader("Members Status")
    ms_display = st.session_state.members_status.copy()
    # progress
    total = len(ms_display)
    filled = (ms_display["Status"] == "✅ Filled").sum()
    progress_value = filled / total if total > 0 else 0
    st.progress(progress_value)
    st.write(f"✅ {filled} of {total} members have submitted the form")
    st.dataframe(ms_display.reset_index(drop=True))
else:
    st.write("No members data available.")

# Download buttons
col1, col2 = st.columns(2)
with col1:
    if not submissions_df.empty:
        buf = BytesIO()
        submissions_df.to_excel(buf, index=Fal)
        buf.seek(0)
        st.download_button("⬇️ Download Submissions Excel", data=buf, file_name="submissions.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with col2:
    if not st.session_state.members_status.empty:
        buf2 = BytesIO()
        st.session_state.members_status.to_excel(buf2, index=False)
        buf2.seek(0)
        st.download_button("⬇️ Download Members Status Excel", data=buf2, file_name="members_status.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Notes: 1) This app does NOT use WhatsApp API — it prepares WhatsApp links for manual sending. 2) 'Open All' may be blocked by browser popup blockers; use individual 'Send' links if that happens.")
