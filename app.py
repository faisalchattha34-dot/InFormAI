import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re

st.set_page_config(page_title="Excel Smart Form + WhatsApp", page_icon="📄", layout="centered")
st.title("📄 Excel Smart Form + WhatsApp Member System")

# --- Step 1: Upload Members File ---
st.header("👥 Upload Members List")
members_file = st.file_uploader("📂 Upload Members Excel File", type=["xlsx"], key="members")

members_df = None
if members_file:
    members_df = pd.read_excel(members_file)
    members_df.columns = [c.strip().title() for c in members_df.columns]
    if "Name" not in members_df.columns or "Whatsapp" not in members_df.columns:
        st.error("❌ Members file must contain columns: 'Name' and 'Whatsapp'")
        members_df = None
    else:
        st.success("✅ Members file loaded successfully!")
        st.dataframe(members_df)

# --- Step 2: Upload Form File ---
st.header("🧾 Upload Form Template File")
form_file = st.file_uploader("📂 Upload Form Excel File", type=["xlsx"], key="form")

if form_file:
    try:
        # Detect header row
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
        st.write("**Detected Columns:**", list(df.columns))

        # Detect dropdowns
        form_file.seek(0)
        wb = load_workbook(form_file, data_only=True)
        ws = wb.active
        dropdown_dict = {}

        if ws.data_validations is not None:
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

        if dropdown_dict:
            st.info("🎯 **Detected Dropdown Columns:**")
            st.table(pd.DataFrame([
                {"Column": col, "Dropdown Options": ", ".join(vals)}
                for col, vals in dropdown_dict.items()
            ]))
        else:
            st.warning("⚠️ No dropdown lists detected in this Excel file.")

        # --- Step 3: Dynamic Form ---
        st.header("🖊️ Fill the Form")
        data = {}
        for col in df.columns:
            if col in dropdown_dict:
                data[col] = st.selectbox(f"{col}", dropdown_dict[col], key=col)
            else:
                data[col] = st.text_input(f"{col}", key=col)

        if "form_submissions" not in st.session_state:
            st.session_state.form_submissions = pd.DataFrame(columns=df.columns)

        if st.button("✅ Submit Form"):
            new_row = pd.DataFrame([data])
            st.session_state.form_submissions = pd.concat(
                [st.session_state.form_submissions, new_row], ignore_index=True
            )
            st.success("🎉 Form submitted successfully!")

        st.subheader("📋 All Submitted Data")
        st.dataframe(st.session_state.form_submissions)

        # --- Step 4: WhatsApp Message Section ---
        if members_df is not None:
            st.header("📱 WhatsApp Share Links (Manual Send)")

            app_link = st.text_input(
                "🔗 Enter Your Streamlit App Link (same for everyone)",
                placeholder="https://your-streamlit-app.streamlit.app"
            )

            if app_link:
                st.write("👇 Click any link to open WhatsApp chat ready to send:")
                wa_links = []
                for _, row in members_df.iterrows():
                    name = str(row["Name"])
                    number = str(row["Whatsapp"]).replace("+", "").replace(" ", "").replace("-", "")
                    message = (
                        f"Hello {name}! Please fill your form here: {app_link}"
                    )
                    wa_url = f"https://wa.me/{number}?text={message.replace(' ', '%20')}"
                    wa_links.append({"Name": name, "WhatsApp": f"[Send to {name}]({wa_url})"})

                st.markdown(pd.DataFrame(wa_links).to_markdown(index=False), unsafe_allow_html=True)

        # --- Step 5: Submission Progress (if members uploaded) ---
        if members_df is not None and len(st.session_state.form_submissions) > 0:
            filled = len(st.session_state.form_submissions)
            total = len(members_df)
            progress = filled / total if total > 0 else 0
            st.header("📊 Submission Progress")
            st.progress(progress)
            st.write(f"✅ {filled} of {total} members have submitted the form")

        # --- Step 6: Download updated submissions ---
        output = BytesIO()
        st.session_state.form_submissions.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="⬇️ Download Submitted Data (Excel)",
            data=output,
            file_name="form_submissions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing form: {e}")
